#region Description
<#     
       .NOTES
       ===========================================================================
       Created on:         2017/11/03 
       Created by:         Drago Petrovic | Dominic Manning
       Organization:       MSB365.blog
       Filename:           ExchangeSuite.ps1     

       Find us on:
             * Website:         https://www.msb365.blog
             * Technet:         https://social.technet.microsoft.com/Profile/MSB365
             * LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
             * MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
       ===========================================================================
       .DESCRIPTION
             This script helps you to prepare a Windows server 2012 R2 or Windows server 2016 for the Exchange installation, also the Exchange installation by it self
             and the most nessesary POST tasks.
             By running the script it will figure out automaticaly which server OS is currently running and it will present the right options for it.
       
       .NOTES
             

       .COPYRIGHT
       Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
       to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
       and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
       WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
       ===========================================================================
       .CHANGE LOG
             V0.10, 2017/12/27 - Initial version
             V0.20, 2018/01/18 - Bug fixes
             V0.30, 2018/02/12 - Adding Admin requirements
             V0.40, 2018/05/23 - Creating POST EXCHANGE Tasks
             V0.50, 2018/06/17 - Creating Office 365 Tasks
             V0.60, 2018/06/19 - Bug fixes
             V0.70, 2018/10/01 - Bug fixes
             V0.80, 2018/11/23 - Optimizing loop
             V0.90, 2018/12/10 - Modify Prerequisite 
             V1.00, 2024/01/10 - Added Mobile Device Reporting (Script part from Tony Redmond)


--- keep it simple, but significant ---
#>
#endregion
##############################################################################################################
[cmdletbinding()]
param(
[switch]$accepteula,
[switch]$v)

###############################################################################
#Script Name variable
$Scriptname = "ExchangeSuite V1.0"
$RKEY = "MSB365_ExchangeSuite_V10"
###############################################################################

[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

function ShowEULAPopup($mode)
{
    $EULA = New-Object -TypeName System.Windows.Forms.Form
    $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
    $btnAcknowledge = New-Object System.Windows.Forms.Button
    $btnCancel = New-Object System.Windows.Forms.Button

    $EULA.SuspendLayout()
    $EULA.Name = "MIT"
    $EULA.Text = "$Scriptname - License Agreement"

    $richTextBox1.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $richTextBox1.Location = New-Object System.Drawing.Point(12,12)
    $richTextBox1.Name = "richTextBox1"
    $richTextBox1.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $richTextBox1.Size = New-Object System.Drawing.Size(776, 397)
    $richTextBox1.TabIndex = 0
    $richTextBox1.ReadOnly=$True
    $richTextBox1.Add_LinkClicked({Start-Process -FilePath $_.LinkText})
    $richTextBox1.Rtf = @"
{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fswiss\fprq2\fcharset0 Segoe UI;}{\f1\fnil\fcharset0 Calibri;}{\f2\fnil\fcharset0 Microsoft Sans Serif;}}
{\colortbl ;\red0\green0\blue255;}
{\*\generator Riched20 10.0.19041}{\*\mmathPr\mdispDef1\mwrapIndent1440 }\viewkind4\uc1
\pard\widctlpar\f0\fs19\lang1033 MSB365 SOFTWARE MIT LICENSE\par
Copyright (c) 2024 Drago Petrovic\par
$Scriptname \par/
\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}These license terms are an agreement between you and MSB365 (or one of its affiliates). IF YOU COMPLY WITH THESE LICENSE TERMS, YOU HAVE THE RIGHTS BELOW. BY USING THE SOFTWARE, YOU ACCEPT THESE TERMS.\par
\par
MIT License\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}\par
\pard
{\pntext\f0 1.\tab}{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{\pntxta.}}
\fi-360\li360 Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: \par
\pard\widctlpar\par
\pard\widctlpar\li360 The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\par
\par
\pard\widctlpar\fi-360\li360 2.\tab THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 3.\tab IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 4.\tab DISCLAIMER OF WARRANTY. THE SOFTWARE IS PROVIDED \ldblquote AS IS,\rdblquote  WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL MSB365 OR ITS LICENSORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THE SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 5.\tab LIMITATION ON AND EXCLUSION OF DAMAGES. IF YOU HAVE ANY BASIS FOR RECOVERING DAMAGES DESPITE THE PRECEDING DISCLAIMER OF WARRANTY, YOU CAN RECOVER FROM MICROSOFT AND ITS SUPPLIERS ONLY DIRECT DAMAGES UP TO U.S. $1.00. YOU CANNOT RECOVER ANY OTHER DAMAGES, INCLUDING CONSEQUENTIAL, LOST PROFITS, SPECIAL, INDIRECT, OR INCIDENTAL DAMAGES. This limitation applies to (i) anything related to the Software, services, content (including code) on third party Internet sites, or third party applications; and (ii) claims for breach of contract, warranty, guarantee, or condition; strict liability, negligence, or other tort; or any other claim; in each case to the extent permitted by applicable law. It also applies even if MSB365 knew or should have known about the possibility of the damages. The above limitation or exclusion may not apply to you because your state, province, or country may not allow the exclusion or limitation of incidental, consequential, or other damages.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 6.\tab ENTIRE AGREEMENT. This agreement, and any other terms MSB365 may provide for supplements, updates, or third-party applications, is the entire agreement for the software.\par
\pard\widctlpar\qj\par
\pard\widctlpar\fi-360\li360\qj 7.\tab A complete script documentation can be found on the website https://www.msb365.blog.\par
\pard\widctlpar\par
\pard\sa200\sl276\slmult1\f1\fs22\lang9\par
\pard\f2\fs17\lang2057\par
}
"@
    $richTextBox1.BackColor = [System.Drawing.Color]::White
    $btnAcknowledge.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnAcknowledge.Location = New-Object System.Drawing.Point(544, 415)
    $btnAcknowledge.Name = "btnAcknowledge";
    $btnAcknowledge.Size = New-Object System.Drawing.Size(119, 23)
    $btnAcknowledge.TabIndex = 1
    $btnAcknowledge.Text = "Accept"
    $btnAcknowledge.UseVisualStyleBackColor = $True
    $btnAcknowledge.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::Yes})

    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.Location = New-Object System.Drawing.Point(669, 415)
    $btnCancel.Name = "btnCancel"
    $btnCancel.Size = New-Object System.Drawing.Size(119, 23)
    $btnCancel.TabIndex = 2
    if($mode -ne 0)
    {
   $btnCancel.Text = "Close"
    }
    else
    {
   $btnCancel.Text = "Decline"
    }
    $btnCancel.UseVisualStyleBackColor = $True
    $btnCancel.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::No})

    $EULA.AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
    $EULA.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $EULA.ClientSize = New-Object System.Drawing.Size(800, 450)
    $EULA.Controls.Add($btnCancel)
    $EULA.Controls.Add($richTextBox1)
    if($mode -ne 0)
    {
   $EULA.AcceptButton=$btnCancel
    }
    else
    {
        $EULA.Controls.Add($btnAcknowledge)
   $EULA.AcceptButton=$btnAcknowledge
        $EULA.CancelButton=$btnCancel
    }
    $EULA.ResumeLayout($false)
    $EULA.Size = New-Object System.Drawing.Size(800, 650)

    Return ($EULA.ShowDialog())
}

function ShowEULAIfNeeded($toolName, $mode)
{
$eulaRegPath = "HKCU:Software\Microsoft\$RKEY"
$eulaAccepted = "No"
$eulaValue = $toolName + " EULA Accepted"
if(Test-Path $eulaRegPath)
{
$eulaRegKey = Get-Item $eulaRegPath
$eulaAccepted = $eulaRegKey.GetValue($eulaValue, "No")
}
else
{
$eulaRegKey = New-Item $eulaRegPath
}
if($mode -eq 2) # silent accept
{
$eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
else
{
if($eulaAccepted -eq "No")
{
$eulaAccepted = ShowEULAPopup($mode)
if($eulaAccepted -eq [System.Windows.Forms.DialogResult]::Yes)
{
        $eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
}
}
return $eulaAccepted
}

if ($accepteula)
    {
         ShowEULAIfNeeded "DS Authentication Scripts:" 2
         "EULA Accepted"
    }
else
    {
        $eulaAccepted = ShowEULAIfNeeded "DS Authentication Scripts:" 0
        if($eulaAccepted -ne "Yes")
            {
                "EULA Declined"
                exit
            }
         "EULA Accepted"
    }
###############################################################################
#region Admin Prerequisite check 
####################
# Prerequisite check
####################
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
	Write-Host "Administrator priviliges are required. Please restart this script with elevated rights." -ForegroundColor Red
	Pause
	Throw "Administrator priviliges are required. Please restart this script with elevated rights."
}
#endregion
#region Global Variable Definitions
##################################
#   Global Variable Definitions  #
##################################

$Ver = (Get-WMIObject win32_OperatingSystem).Version
$OSCheck = $false
$Choice = "None"
$Date = get-date -Format "MM.dd.yyyy-hh.mm-tt"
$DownloadFolder = "c:\install"
$CurrentPath = (Get-Item -Path ".\" -Verbose).FullName
$Reboot = $false
$Error.clear()
Start-Transcript -path "$CurrenPath\$date-Set-Prerequisites.txt" | Out-Null
Clear-Host
# Pushd
#endregion
#region Global Functions
############################################################
#   Global Functions - Shared between 2012 (R2) and 2016   #
############################################################

# Begin BITSCheck function
function BITSCheck
{
	$Bits = Get-Module BitsTransfer
	if ($Bits -eq $null)
	{
		write-host "Importing the BITS module." -ForegroundColor cyan
		try
		{
			Import-Module BitsTransfer -erroraction STOP
		}
		catch
		{
			write-host "Server Management module could not be loaded." -ForegroundColor Red
		}
	}
} # End BITSCheck function

# Begin ModuleStatus function
function ModuleStatus
{
	$module = Get-Module -name "ServerManager" -erroraction STOP
	
	if ($module -eq $null)
	{
		try
		{
			Import-Module -Name "ServerManager" -erroraction STOP
			# return $null
		}
		catch
		{
			write-host " "; write-host "Server Manager module could not be loaded." -ForegroundColor Red
		}
	}
	else
	{
		# write-host "Server Manager module is already imported." -ForegroundColor Cyan
		# return $null
	}
	write-host " "
} # End ModuleStatus function

# Begin FileDownload function
function FileDownload
{
	param ($sourcefile)
	$Internetaccess = (Get-NetConnectionProfile -IPv4Connectivity Internet).ipv4connectivity
	If ($Internetaccess -eq "Internet")
	{
		if (Test-path $DownloadFolder)
		{
			write-host "Target folder $DownloadFolder exists." -foregroundcolor white
		}
		else
		{
			New-Item $DownloadFolder -type Directory | Out-Null
		}
		BITSCheck
		[string]$DownloadFile = $sourcefile.Substring($sourcefile.LastIndexOf("/") + 1)
		if (Test-Path "$DownloadFolder\$DownloadFile")
		{
			write-host "The file $DownloadFile already exists in the $DownloadFolder folder." -ForegroundColor Cyan
		}
		else
		{
			Start-BitsTransfer -Source "$SourceFile" -Destination "$DownloadFolder\$DownloadFile"
		}
	}
	else
	{
		write-host "This machine does not have internet access and thus cannot download required files. Please resolve!" -ForegroundColor Red
	}
} # End FileDownload function
#endregion
#region High Performance power plan
# Configure the Server for the High Performance power plan
function highperformance
{
	write-host " "
	$HighPerf = powercfg -l | %{ if ($_.contains("High performance")) { $_.split()[3] } }
	$CurrPlan = $(powercfg -getactivescheme).split()[3]
	if ($CurrPlan -ne $HighPerf)
	{
		powercfg -setactive $HighPerf
		CheckPowerPlan
	}
	else
	{
		if ($CurrPlan -eq $HighPerf)
		{
			write-host " "; write-host "The power plan is already set to " -nonewline; write-host "High Performance." -foregroundcolor green; write-host " "
		}
	}
}
#endregion
#region server power management
# Check the server power management
function CheckPowerPlan
{
	$HighPerf = powercfg -l | %{ if ($_.contains("High performance")) { $_.split()[3] } }
	$CurrPlan = $(powercfg -getactivescheme).split()[3]
	if ($CurrPlan -eq $HighPerf)
	{
		write-host " "; write-host "The power plan now is set to " -nonewline; write-host "High Performance." -foregroundcolor green; write-host " "
	}
}
#endregion
#region NIC power management
# Turn off NIC power management
function PowerMgmt
{
	write-host " "
	$NICs = Get-WmiObject -Class Win32_NetworkAdapter | Where-Object{ $_.PNPDeviceID -notlike "ROOT\*" -and $_.Manufacturer -ne "Microsoft" -and $_.ConfigManagerErrorCode -eq 0 -and $_.ConfigManagerErrorCode -ne 22 }
	Foreach ($NIC in $NICs)
	{
		$NICName = $NIC.Name
		$DeviceID = $NIC.DeviceID
		If ([Int32]$DeviceID -lt 10)
		{
			$DeviceNumber = "000" + $DeviceID
		}
		Else
		{
			$DeviceNumber = "00" + $DeviceID
		}
		$KeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\$DeviceNumber"
		
		If (Test-Path -Path $KeyPath)
		{
			$PnPCapabilities = (Get-ItemProperty -Path $KeyPath).PnPCapabilities
			# Check to see if the value is 24 and if not, set it to 24
			If ($PnPCapabilities -ne 24) { Set-ItemProperty -Path $KeyPath -Name "PnPCapabilities" -Value 24 | Out-Null }
			# Verify the value is now set to or was set to 24
			If ($PnPCapabilities -eq 24) { write-host " "; write-host "Power Management has already been " -NoNewline; write-host "disabled" -ForegroundColor Green; write-host " " }
		}
	}
}
#endregion
#region RC4
# Disable RC4
function DisableRC4
{
	write-host " "
	# Define Registry keys to look for
	$base = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\" -erroraction silentlycontinue
	$val1 = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 128/128\" -erroraction silentlycontinue
	$val2 = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 40/128\" -erroraction silentlycontinue
	$val3 = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 56/128\" -erroraction silentlycontinue
	
	# Define Values to add
	$registryBase = "Ciphers"
	$registryPath1 = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 128/128\"
	$registryPath2 = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 40/128\"
	$registryPath3 = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 56/128\"
	$Name = "Enabled"
	$value = "0"
	$ssl = 0
	$checkval1 = Get-Itemproperty -Path "$registrypath1" -name $name -erroraction silentlycontinue
	$checkval2 = Get-Itemproperty -Path "$registrypath2" -name $name -erroraction silentlycontinue
	$checkval3 = Get-Itemproperty -Path "$registrypath3" -name $name -erroraction silentlycontinue
	
	# Formatting for output
	write-host " "
	
	# Add missing registry keys as needed
	If ($base -eq $null)
	{
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL", $true)
		$key.CreateSubKey('Ciphers')
		$key.Close()
	}
	else
	{
		write-host "The " -nonewline; write-host "Ciphers" -ForegroundColor green -NoNewline; write-host " Registry key already exists."
	}
	
	If ($val1 -eq $null)
	{
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers", $true)
		$key.CreateSubKey('RC4 128/128')
		$key.Close()
	}
	else
	{
		write-host "The " -nonewline; write-host "Ciphers\RC4 128/128" -ForegroundColor green -NoNewline; write-host " Registry key already exists."
	}
	
	If ($val2 -eq $null)
	{
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers", $true)
		$key.CreateSubKey('RC4 40/128')
		$key.Close()
		New-ItemProperty -Path $registryPath2 -Name $name -Value $value
	}
	else
	{
		write-host "The " -nonewline; write-host "Ciphers\RC4 40/128" -ForegroundColor green -NoNewline; write-host " Registry key already exists."
	}
	
	If ($val3 -eq $null)
	{
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers", $true)
		$key.CreateSubKey('RC4 56/128')
		$key.Close()
	}
	else
	{
		write-host "The " -nonewline; write-host "Ciphers\RC4 56/128" -ForegroundColor green -NoNewline; write-host " Registry key already exists."
	}
	
	# Add the enabled value to disable RC4 Encryption
	If ($checkval1.enabled -ne "0")
	{
		try
		{
			New-ItemProperty -Path $registryPath1 -Name $name -Value $value -force; $ssl++
		}
		catch
		{
			$SSL--
		}
	}
	else
	{
		write-host "The registry value " -nonewline; write-host "Enabled" -ForegroundColor green -NoNewline; write-host " exists under the RC4 128/128 Registry Key."; $ssl++
	}
	If ($checkval2.enabled -ne "0")
	{
		write-host $checkval2
		try
		{
			New-ItemProperty -Path $registryPath2 -Name $name -Value $value -force; $ssl++
		}
		catch
		{
			$SSL--
		}
	}
	else
	{
		write-host "The registry value " -nonewline; write-host "Enabled" -ForegroundColor green -NoNewline; write-host " exists under the RC4 40/128 Registry Key."; $ssl++
	}
	If ($checkval3.enabled -ne "0")
	{
		try
		{
			New-ItemProperty -Path $registryPath3 -Name $name -Value $value -force; $ssl++
		}
		catch
		{
			$SSL--
		}
	}
	else
	{
		write-host "The registry value " -nonewline; write-host "Enabled" -ForegroundColor green -NoNewline; write-host " exists under the RC4 56/128 Registry Key."; $ssl++
	}
	
	# SSL Check totals
	If ($ssl -eq "3")
	{
		write-host " "; write-host "RC4 " -ForegroundColor yellow -NoNewline; write-host "is completely disabled on this server."; write-host " "
	}
	If ($ssl -lt "3")
	{
		write-host " "; write-host "RC4 " -ForegroundColor yellow -NoNewline; write-host "only has $ssl part(s) of 3 disabled.  Please check the registry to manually to add these values"; write-host " "
	}
} # End of Disable RC4 function
#endregion
#region SSL
# Disable SSL 3.0
function DisableSSL3
{
	write-host " "
	$TestPath1 = Get-Item -Path "HKLM:\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0" -erroraction silentlycontinue
	$TestPath2 = Get-Item -Path "HKLM:\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server" -erroraction silentlycontinue
	$registrypath = "HKLM:\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server"
	$Name = "Enabled"
	$value = "0"
	$checkval1 = Get-Itemproperty -Path "$registrypath" -name $name -erroraction silentlycontinue
	
	# Check for SSL 3.0 Reg Key
	If ($TestPath1 -eq $null)
	{
		$key = (get-item HKLM:\).OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols", $true)
		$key.CreateSubKey('SSL 3.0')
		$key.Close()
	}
	else
	{
		write-host "The " -nonewline; write-host "SSL 3.0" -ForegroundColor green -NoNewline; write-host " Registry key already exists."
	}
	
	# Check for SSL 3.0\Server Reg Key
	If ($TestPath2 -eq $null)
	{
		$key = (get-item HKLM:\).OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0", $true)
		$key.CreateSubKey('Server')
		$key.Close()
	}
	else
	{
		write-host "The " -nonewline; write-host "SSL 3.0\Servers" -ForegroundColor green -NoNewline; write-host " Registry key already exists."
	}
	
	# Add the enabled value to disable SSL 3.0 Support
	If ($checkval1.enabled -ne "0")
	{
		try
		{
			New-ItemProperty -Path $registryPath -Name $name -Value $value -force; $ssl++
		}
		catch
		{
			$SSL--
		}
	}
	else
	{
		write-host "The registry value " -nonewline; write-host "Enabled" -ForegroundColor green -NoNewline; write-host " exists under the SSL 3.0\Server Registry Key."
	}
} # End of Disable SSL 3.0 function
#endregion
#region Unified Communications
# Function - Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
function Install-WinUniComm4
{
	write-host " "
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
	if ($val.DisplayVersion -ne "5.0.8308.0")
	{
		if ($val.DisplayVersion -ne "5.0.8132.0")
		{
			if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A41CBE7D-949C-41DD-9869-ABBD99D753DA}") -eq $false)
			{
				Write-Host "`nMicrosoft Unified Communications Managed API 4.0 is not installed.  Downloading and installing now." -foregroundcolor yellow
				Install-NewWinUniComm4
			}
			else
			{
				Write-Host "`nAn old version of Microsoft Unified Communications Managed API 4.0 is installed."
				UnInstall-WinUniComm4
				Write-Host "`nMicrosoft Unified Communications Managed API 4.0 has been uninstalled.  Downloading and installing now." -foregroundcolor green
				Install-NewWinUniComm4
			}
		}
		else
		{
			Write-Host "`nThe Preview version of Microsoft Unified Communications Managed API 4.0 is installed."
			UnInstall-WinUniComm4
			Write-Host "`nMicrosoft Unified Communications Managed API 4.0 has been uninstalled.  Downloading and installing now." -foregroundcolor green
			Install-NewWinUniComm4
		}
	}
	else
	{
		write-host "The correct version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
		write-host "installed." -ForegroundColor green
	}
	write-host " "
} # end Install-WinUniComm4
#endregion
#region Install Unified Communications
# Install Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
function Install-NewWinUniComm4
{
	FileDownload "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe"
	Set-Location $DownloadFolder
	[string]$expression = ".\UcmaRuntimeSetup.exe /quiet /norestart /l* $targetfolder\WinUniComm4.log"
	Write-Host "File: UcmaRuntimeSetup.exe installing..." -NoNewLine
	Invoke-Expression $expression
	Start-Sleep -Seconds 20
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
	if ($val.DisplayVersion -ne "5.0.8308.0")
	{
		Write-Host "`nMicrosoft Unified Communications Managed API 4.0 is now installed" -Foregroundcolor Green
	}
	write-host " "
} # end Install-NewWinUniComm4
#endregion
#region PageFile
# Configure PageFile for Exchange
function ConfigurePageFile
{
	$Stop = $False
	$WMIQuery = $False
	
	# Remove Existing PageFile
	try
	{
		Set-CimInstance -Query "Select * from win32_computersystem" -Property @{ automaticmanagedpagefile = "False" }
	}
	catch
	{
		write-host "Cannot remove the existing pagefile." -ForegroundColor Red
		$WMIQuery = $True
	}
	# Remove PageFile with WMI if CIM fails
	If ($WMIQuery)
	{
		Try
		{
			$CurrentPageFile = Get-WmiObject -Class Win32_PageFileSetting
			$name = $CurrentPageFile.Name
			$CurrentPageFile.delete()
		}
		catch
		{
			write-host "The server $server cannot be reached via CIM or WMI." -ForegroundColor Red
			$Stop = $True
		}
	}
	
	# Get RAM and set ideal PageFileSize
	$GB = 1048576
	
	try
	{
		$RamInMb = (Get-CIMInstance -computername $name -Classname win32_physicalmemory -ErrorAction Stop | measure-object -property capacity -sum).sum/$GB
		$ExchangeRAM = $RAMinMb + 10
		# Set maximum pagefile size to 32 GB + 10 MB
		if ($ExchangeRAM -gt 32778) { $ExchangeRAM = 32778 }
	}
	catch
	{
		write-host "Cannot acquire the amount of RAM in the server." -ForegroundColor Red
		$stop = $true
	}
	# Get RAM and set ideal PageFileSize - WMI Method
	If ($WMIQuery)
	{
		Try
		{
			$RamInMb = (Get-wmiobject -computername $server -Classname win32_physicalmemory -ErrorAction Stop | measure-object -property capacity -sum).sum/$GB
			$ExchangeRAM = $RAMinMb + 10
			
			# Set maximum pagefile size to 32 GB + 10 MB
			if ($ExchangeRAM -gt 32778) { $ExchangeRAM = 32778 }
		}
		catch
		{
			write-host "Cannot acquire the amount of RAM in the server with CIM or WMI queries." -ForegroundColor Red
			$stop = $true
		}
	}
	
	# Reset WMIQuery
	$WMIQuery = $False
	
	if ($stop -ne $true)
	{
		# Configure PageFile
		try
		{
			Set-CimInstance -Query "Select * from win32_PageFileSetting" -Property @{ InitialSize = $ExchangeRAM; MaximumSize = $ExchangeRAM }
		}
		catch
		{
			write-host "Cannot configure the PageFile correctly." -ForegroundColor Red
		}
		If ($WMIQuery)
		{
			Try
			{
				Set-WMIInstance -computername $server -class win32_PageFileSetting -arguments @{ name = "$name"; InitialSize = $ExchangeRAM; MaximumSize = $ExchangeRAM }
			}
			catch
			{
				write-host "Cannot configure the PageFile correctly." -ForegroundColor Red
				$stop = $true
			}
		}
		if ($stop -ne $true)
		{
			$pagefile = Get-CimInstance win32_PageFileSetting -Property * | select-object Name, initialsize, maximumsize
			$name = $pagefile.name; $max = $pagefile.maximumsize; $min = $pagefile.initialsize
			write-host " "
			write-host "This server's pagefile, located at " -ForegroundColor white -NoNewline
			write-host "$name" -ForegroundColor green -NoNewline
			write-host ", is now configured for an initial size of " -ForegroundColor white -NoNewline
			write-host "$min MB " -ForegroundColor green -NoNewline
			write-host "and a maximum size of " -ForegroundColor white -NoNewline
			write-host "$max MB." -ForegroundColor Green
			write-host " "
		}
		else
		{
			write-host "The PageFile cannot be configured at this time." -ForegroundColor Red
		}
	}
	else
	{
		write-host "The PageFile cannot be configured at this time." -ForegroundColor Red
	}
}
#endregion
#region ----- Menu 2012 -----
######################################################
#    This section is for the Windows 2012 (R2) OS    #
######################################################

function Code2012
{
	
	# Start code block for Windows 2012 or 2012 R2
	
	$Menu2012 = {
		
		write-host " ********************************************************************" -ForegroundColor Cyan
		write-host " Exchange Server 2016 [On Windows 2012 (R2)]" -ForegroundColor Cyan
		Write-Host "               --- keep it simple, but significant ---" -ForegroundColor Gray
		Write-Host " >>> MSB365 2018 Suite - www.msb365.blog <<<" -ForegroundColor Cyan
		write-host " ********************************************************************" -ForegroundColor Cyan
		write-host " "
		write-host " Please select an option from the list below:" -ForegroundColor White
		write-host " "
		write-host " EXCHANGE SETUP PREREQUISITES (* Exchange media required!)" -ForegroundColor Yellow
		write-host " ---------------------------------------------------------" -foregroundcolor yellow
		write-host " "
		write-host " 1) Launch Windows Update" -ForegroundColor White
		write-host " 2) Check Prerequisites for Mailbox role | Multirole" -ForegroundColor White
		write-host " 3) Check Prerequisites for Edge role" -ForegroundColor White
		write-host " "
		write-host " 4) Install Mailbox prerequisites - Part 1 - RTM and CU1 [.NET 4.5.2]" -ForegroundColor white
		write-host " 5) Install Mailbox prerequisites - Part 1 - CU2 + [.NET 4.6.1]" -ForegroundColor white
		write-host " 6) Install Mailbox prerequisites - Part 2 - All Versions" -ForegroundColor white
		write-host " 7) Install Edge Transport Server prerequisites - RTM and CU1 [.NET 4.5.2]" -ForegroundColor white
		write-host " 8) Install Edge Transport Server prerequisites - CU2 + [.NET 4.6.1]" -ForegroundColor white
		write-host " "
		write-host " 9) Install - One-Off - .NET 4.5.2 [MBX or Edge] - Exchange 2016 RTM or CU1" -ForegroundColor white
		write-host " 10) Install - One-Off - .NET 4.6.1 [MBX or Edge] - Exchange 2016 CU2+" -ForegroundColor white
		write-host " 11) Install - One-Off - Windows Features [MBX]" -ForegroundColor white
		write-host " 12) Install - One Off - Unified Communications Managed API 4.0" -ForegroundColor white
		write-host " "
		write-host " 13) Prepare Schema *" -ForegroundColor White
		write-host " 14) Prepare Active Directory and Domains *" -ForegroundColor White
		write-host " "
		write-host " 15) Set Power Plan to High Performance" -ForegroundColor white
		write-host " 16) Disable Power Management for NICs." -ForegroundColor white
		write-host " 17) Disable SSL 3.0 Support" -ForegroundColor white
		write-host " 18) Disable RC4 Support" -ForegroundColor white
		write-host " "
		write-host " "
		Write-Host " EXCHANGE SETUP TASKS (* Exchange media required!) " -ForegroundColor yellow
		write-host " ---------------------------------------------------------" -foregroundcolor yellow
		write-host " "
		Write-Host " 30) Start Exchange Server setup *" -ForegroundColor Magenta
		write-host " "
		write-host " "
		write-host " POST EXCHANGE 2016 INSTALL" -ForegroundColor Yellow
		write-host " ---------------------------------------------------------" -foregroundcolor yellow
		write-host " "
		write-host " 40) Configure PageFile to RAM + 10 MB" -ForegroundColor green
		Write-Host " 41) Show Exchange URI" -ForegroundColor white
		Write-Host " 42) Configure Exchange URLs" -ForegroundColor white
		Write-Host " 43) Disable UAC" -ForegroundColor white
		Write-Host " 44) Disable Windows Firewall" -ForegroundColor white
		write-host " "
		Write-Host " 45) Create receive connector" -ForegroundColor White
		write-host " 46) Create send connector" -ForegroundColor White
		write-host " 47) Create DAG" -ForegroundColor white
		#write-host " 48) -Create Exchange Hybrid mode" -ForegroundColor Magenta
		write-host " "
		write-host " 49) Create Certificate request" -ForegroundColor White
		Write-Host " 50) set mailaddress policies" -ForegroundColor White
		write-host " "
		write-host " 51) Enable UM for all Mailboxes" -ForegroundColor White
		write-host " 52) Remove  old EAS devices" -ForegroundColor White
		#write-host " 53) -Deploy Microsoft Teams Desktop Client" -ForegroundColor White
		write-host " "
		#write-host " 54) -Order certificate >>GO DADDY<<" -ForegroundColor White
		#write-host " 55) -Order certificate >>DIGICERT<<" -ForegroundColor White
		write-host " "
		write-host " "
		write-host " OPERATING EXCHANGE" -ForegroundColor Yellow
		write-host " ---------------------------------------------------------" -foregroundcolor yellow
		write-host " "
		write-host " 60) Generate Health Report for an Exchange Server 2016/2013/2010 Environment" -ForegroundColor white
		write-host " 61) Generate Exchange Environment Reports" -ForegroundColor white
		#write-host " 62) -Generate Mailbox Size and Information Reports" -ForegroundColor white
		write-host " 63) Generate Reports for Exchange ActiveSync Device Statistics" -ForegroundColor white
		#write-host " 64) -Exchange Analyzer" -ForegroundColor white
		write-host " 65) Generate Report Total Emails Sent and Received Per Day and Size" -ForegroundColor white
		write-host " 66) Generate HTML Report for Mailbox Permissions" -ForegroundColor white
		write-host " "
		write-host " "
		write-host " Microsoft 365 Operation (Exchange online)" -ForegroundColor Yellow
		write-host " ---------------------------------------------------------" -foregroundcolor yellow
		write-host " "
		write-host " 70) Export Office 365 User Last Logon Date to CSV File" -ForegroundColor white
		write-host " 71) List all Distribution Groups and their Membership in Office 365" -ForegroundColor white
		write-host " 72) Office 365 Mail Traffic Statistics by User" -ForegroundColor white
		write-host " 73) Export a Licence reconciliation report from Office 365" -ForegroundColor white
		write-host " 74) Export mailbox permissions from Office 365 to CSV file" -ForegroundColor white
		write-host " 75) Microsoft 365 Mailboxes with Synchronized Mobile Devices" -ForegroundColor white
		#write-host " 75) Set Calendar Permission in Office 365 Exchange Online" -ForegroundColor white
		write-host " "
		write-host " "
		write-host " OPERATING EXCHANGE" -ForegroundColor Yellow
		write-host " ---------------------------------------------------------" -foregroundcolor yellow
		write-host " "
		write-host " 98) Restart the Server" -foregroundcolor red
		write-host " 99) Exit" -foregroundcolor cyan
		write-host " "
		write-host " Select an option.. [1-99]? " -foregroundcolor white -nonewline
	}
	#endregion
	#region .NET 4.5.2
	# Function - .NET 4.5.2
	function Install-DotNET452
	{
		# .NET 4.5.2
		$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name "Release"
		if ($val.Release -lt "379893")
		{
			FileDownload "http://download.microsoft.com/download/E/2/1/E21644B5-2DF2-47C2-91BD-63C560427900/NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
			Set-Location $DownloadFolder
			[string]$expression = ".\NDP452-KB2901907-x86-x64-AllOS-ENU.exe /quiet /norestart /l* $DownloadFolder\DotNET452.log"
			Write-Host "File: NDP452-KB2901907-x86-x64-AllOS-ENU.exe installing..." -NoNewLine
			Invoke-Expression $expression
			Start-Sleep -Seconds 20
			Write-Host "`n.NET 4.5.2 is now installed" -Foregroundcolor Green
		}
		else
		{
			Write-Host "`n.NET 4.5.2 already installed" -Foregroundcolor Green
		}
	} # end Install-DotNET452
	#endregion
	#region .NET 4.6.1
	# Function - .NET 4.6.1
	function Install-DotNET461
	{
		
		# .NET 4.6.1 install function
		function NET461
		{
			FileDownload "https://download.microsoft.com/download/E/4/1/E4173890-A24A-4936-9FC9-AF930FE3FA40/NDP461-KB3102436-x86-x64-AllOS-ENU.exe"
			Set-Location $DownloadFolder
			[string]$expression = ".\NDP461-KB3102436-x86-x64-AllOS-ENU.exe /quiet /norestart /l* $DownloadFolder\DotNET461.log"
			write-host " "
			Write-Host "File: NDP461-KB3102436-x86-x64-AllOS-ENU.exe installing..." -NoNewLine
			Invoke-Expression $expression
			Start-Sleep -Seconds 60
			Write-Host "`n.NET 4.6.1 is now installed" -Foregroundcolor Green
			write-host " "
			$Reboot = $true
		}
		
		# .NET 4.6.1 Post Install Hotfix
		function NET461-HotFix
		{
			
			if ((Get-WMIObject win32_OperatingSystem).Version -match '6.2')
			{
				$nethotfixcheck = get-hotfix | where { $_.HotFixid -eq "KB3146714" }
				if ($nethotfixcheck -eq $null)
				{
					Write-Host "The hotfix for NET 4.6.1 is not installed" -Foregroundcolor yellow
					# Download the Hotfix
					FileDownload "http://download.microsoft.com/download/E/F/1/EF1FB34B-58CB-4568-85EC-FA359387E328/Windows8-RT-KB3146714-x64.msu"
					Set-Location $DownloadFolder
					[string]$expression = "wusa.exe .\Windows8-RT-KB3146714-x64.msu /quiet /norestart"
					write-host " "
					Write-Host "File: Windows8-RT-KB3146714-x64.msu installing..." -NoNewLine
					Invoke-Expression $expression
					Start-Sleep -Seconds 60
					Write-Host "`n.HotFix KB3146714 is now installed" -Foregroundcolor Green
					write-host " "
				}
				else
				{
					write-host "The hotfix for .NET 4.6.1 is installed." -foregroundcolor cyan
				}
			}
			
			if ((Get-WMIObject win32_OperatingSystem).Version -match '6.3')
			{
				$nethotfixcheck = get-hotfix | where { $_.HotFixid -eq "KB3146715" }
				if ($nethotfixcheck -eq $null)
				{
					Write-Host "The hotfix for NET 4.6.1 is not installed" -Foregroundcolor yellow
					
					# Download the Hotfix
					FileDownload "http://download.microsoft.com/download/E/F/1/EF1FB34B-58CB-4568-85EC-FA359387E328/Windows8.1-KB3146715-x64.msu"
					Set-Location $DownloadFolder
					[string]$expression = "wusa.exe .\Windows8.1-KB3146715-x64.msu /quiet /norestart"
					write-host " "
					Write-Host "File: Windows8.1-KB3146715-x64.msu installing..." -NoNewLine
					Invoke-Expression $expression
					Start-Sleep -Seconds 60
					Write-Host "`n.HotFix KB3146715 is now installed" -Foregroundcolor Green
					write-host " "
				}
				else
				{
					write-host "The hotfix for .NET 4.6.1 is installed." -foregroundcolor cyan
				}
			}
		}
		
		# Check for .NET 4.6.1    
		$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name "Release"
		if ($val.Release -lt "394271")
		{
			write-host " "
			write-host ".NET 4.6.1 is not installed." -ForegroundColor Yellow
			write-host "Checking for hotfix KB2919355 before installing .NET 4.6.1 as it is required in order to install .NET 4.6.1." -foregroundcolor white
			write-host " "
			
			# Check for Hotfix 2919355
			$hotfix = get-hotfix | where { $_.HotFixid -eq "kb2919355" }
			
			#install hotfix if missing or .NET 4.6.1 if the hotfix is installed
			if ($hotfix -ne $null)
			{
				# Install .NET 4.6.1
				write-host "The proper hotfix KB2919355 is installed, " -ForegroundColor white -NoNewline
				write-host "proceeding to install .NET 4.6.1...." -ForegroundColor Yellow
				NET461
			}
			else
			{
				write-host "Hotfix 2919355 is missing, downloading and installing it now." -ForegroundColor Red
				FileDownload "https://download.microsoft.com/download/2/5/6/256CCCFB-5341-4A8D-A277-8A81B21A1E35/Windows8.1-KB2919355-x64.msu"
				Set-Location $DownloadFolder
				[string]$expression = ".\Windows8.1-KB2919355-x64.msu"
				Write-Host "File: Windows8.1-KB2919355-x64.msu installing..." -NoNewLine
				Invoke-Expression $expression
				
				write-host " "; write-host "This update can take an unusually long time to install...." -ForegroundColor Yellow; write-host " "
				Write-Host "`n.Once the HotFix 2919355 is installed, please make sure to reboot your server." -Foregroundcolor Green
				write-host "Once your server has rebooted, make sure to run option 21 once more to install .NET 4.6.1"
				start-sleep 30
			}
		}
		else
		{
			Write-Host "`n.NET 4.6.1 already installed" -Foregroundcolor Green
		}
		
		# Check the .NET 4.6.1 value again (pre - hotfix install)
		$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name "Release"
		# Check for Hotfix
		if ($val.Release -eq "394271")
		{
			
			write-host " "
			write-host "Checking for hotfix required for post .NET 4.6.1 installation." -foregroundcolor yellow
			write-host " "
			
			# Query for and installation of Hotfix for .NET 4.6.1
			NET461-HotFix
		}
	} # end Install-DotNET452
	
	# Function - Check Dot Net Version
	function Check-DotNetVersion
	{
		# .NET 4.5.2 or 4.6.1
		$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name "Release"
		if ($val.Release -lt "379893")
		{
			write-host ".NET 4.5.2 is " -nonewline
			write-host "not installed!" -ForegroundColor red -nonewline
			write-host " - this does not meet the minimum requirements for Exchange to be installed." -ForegroundColor white
			write-host " "
		}
		
		if ($val.Release -eq "379893")
		{
			write-host ".NET 4.5.2 is " -nonewline -foregroundcolor white
			write-host "installed." -ForegroundColor green -NoNewline
			write-host " - this is sufficient for any version of Exchange Server 2016." -ForegroundColor white
			write-host " "
		}
		
		if ($val.Release -eq "394271")
		{
			write-host ".NET 4.6.1 is " -nonewline -foregroundcolor white
			write-host "installed." -ForegroundColor green -nonewline
			write-host " - This version of .NET is suitable for " -NoNewline -foregroundcolor white
			write-host "Exchange Server 2016 CU2 +" -foregroundcolor yellow
			write-host " "
		}
		
	} # End Check Dot Net Version Function
	#endregion
	#region Mailbox Role
	# Mailbox Role - Windows Feature requirements
	function check-MBXprereq
	{
		write-host " "
		write-host "Checking all requirements for the Mailbox Role in Exchange Server 2016....." -foregroundcolor yellow
		write-host " "
		start-sleep 2
		
		# .NET Check
		Check-DotNetVersion
		
		# Windows Feature Check
		$values = @("AS-HTTP-Activation", "Desktop-Experience", "NET-Framework-45-Features", "RPC-over-HTTP-proxy", "RSAT-Clustering", "RSAT-Clustering-CmdInterface", "RSAT-Clustering-Mgmt", "RSAT-Clustering-PowerShell", "Web-Mgmt-Console", "WAS-Process-Model", "Web-Asp-Net45", "Web-Basic-Auth", "Web-Client-Auth", "Web-Digest-Auth", "Web-Dir-Browsing", "Web-Dyn-Compression", "Web-Http-Errors", "Web-Http-Logging", "Web-Http-Redirect", "Web-Http-Tracing", "Web-ISAPI-Ext", "Web-ISAPI-Filter", "Web-Lgcy-Mgmt-Console", "Web-Metabase", "Web-Mgmt-Console", "Web-Mgmt-Service", "Web-Net-Ext45", "Web-Request-Monitor", "Web-Server", "Web-Stat-Compression", "Web-Static-Content", "Web-Windows-Auth", "Web-WMI", "Windows-Identity-Foundation")
		foreach ($item in $values)
		{
			$val = get-Windowsfeature $item
			If ($val.installed -eq $true)
			{
				write-host "The Windows Feature"$item" is " -nonewline
				write-host "installed." -ForegroundColor green
			}
			else
			{
				write-host "The Windows Feature"$item" is " -nonewline
				write-host "not installed!" -ForegroundColor red
			}
		}
		
		# Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit 
		$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
		if ($val.DisplayVersion -ne "5.0.8308.0")
		{
			if ($val.DisplayVersion -ne "5.0.8132.0")
			{
				if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A41CBE7D-949C-41DD-9869-ABBD99D753DA}") -eq $false)
				{
					write-host "No version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
					write-host "not installed!" -ForegroundColor red
					write-host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992."
				}
				else
				{
					write-host "The Preview version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
					write-host "installed." -ForegroundColor red
					write-host "This is the incorrect version of UCMA. " -nonewline -ForegroundColor red
					write-host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992."
				}
			}
			else
			{
				write-host "The wrong version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
				write-host "installed." -ForegroundColor red
				write-host "This is the incorrect version of UCMA. " -nonewline -ForegroundColor red
				write-host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992."
			}
		}
		else
		{
			write-host "The correct version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
			write-host "installed." -ForegroundColor green
		}
	} # End function check-MBXprereq
	#endregion
	#region Edge Transport
	# Edge Transport requirement check
	function check-EdgePrereq
	{
		
		write-host " "
		write-host "Checking all requirements for the Edge Transport Role in Exchange Server 2016....." -foregroundcolor yellow
		write-host " "
		start-sleep 2
		
		# Check .NET version
		Check-DotNetVersion
		
		# Windows Feature AD LightWeight Services
		$values = @("ADLDS")
		foreach ($item in $values)
		{
			$val = get-Windowsfeature $item
			If ($val.installed -eq $true)
			{
				write-host "The Windows Feature"$item" is " -nonewline
				write-host "installed." -ForegroundColor green
				write-host " "
			}
			else
			{
				write-host "The Windows Feature"$item" is " -nonewline
				write-host "not installed!" -ForegroundColor red
				write-host " "
			}
		}
		write-host " "
	} # End Check-EdgePrereq
	#endregion
	#region Install Microsoft Unified Communications
	# Install Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
	function Install-NewWinUniComm4
	{
		FileDownload "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe"
		Set-Location $DownloadFolder
		[string]$expression = ".\UcmaRuntimeSetup.exe /quiet /norestart /l* $DownloadFolder\WinUniComm4.log"
		Write-Host "File: UcmaRuntimeSetup.exe installing..." -NoNewLine
		Invoke-Expression $expression
		Start-Sleep -Seconds 20
		$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
		if ($val.DisplayVersion -ne "5.0.8308.0")
		{
			Write-Host "`nMicrosoft Unified Communications Managed API 4.0 is now installed" -Foregroundcolor Green
		}
		write-host " "
	} # end Install-NewWinUniComm4
	#endregion
	#region Reboot task 
	Do
	{
		if ($Reboot -eq $true) { Write-Host "`t`t`t`t`t`t`t`t`t`n`t`t`t`tREBOOT REQUIRED!`t`t`t`n`t`t`t`t`t`t`t`t`t`n`t`tDO NOT INSTALL EXCHANGE BEFORE REBOOTING!`t`t`n`t`t`t`t`t`t`t`t`t" -backgroundcolor red -foregroundcolor black }
		if ($Choice -ne "None") { Write-Host "Last command: "$Choice -foregroundcolor Yellow }
		invoke-command -scriptblock $Menu2012
		$Choice = Read-Host
		#endregion          
		switch ($Choice)
		{
			#region Option  1) Windows Update
			1 {
				#      Windows Update
				Invoke-Expression "$env:windir\system32\wuapp.exe startmenu"
			}
			#endregion
			#region Option  2) Mailbox Requirement Check
			2 {
				#      Mailbox Requirement Check
				check-MBXprereq
			}
			#endregion
			#region Option  3) Edge Transport Requirement Check
			3 {
				#      Edge Transport Requirement Check
				check-EdgePrereq
			}
			#endregion
			#region Option  4) Prep Mailbox Role - Part 1 - RTM and CU1
			4 {
				#      Prep Mailbox Role - Part 1 - RTM and CU1
				ModuleStatus -name ServerManager
				Install-DotNET452
				Install-WindowsFeature RSAT-ADDS
				
				Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
				highperformance
				PowerMgmt
				$Reboot = $true
			}
			#endregion
			#region Option  5) Prep Mailbox Role - Part 1 - CU2 +
			5 {
				#      Prep Mailbox Role - Part 1 - CU2 +
				ModuleStatus -name ServerManager
				Install-DotNET461
				Install-WindowsFeature RSAT-ADDS
				Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
				highperformance
				PowerMgmt
				$Reboot = $true
			}
			#endregion
			#region Option  6) Prep Mailbox Role - Part 2 - All Versions
			6 {
				#      Prep Mailbox Role - Part 2 - All Versions
				ModuleStatus -name ServerManager
				Install-WinUniComm4
				$Reboot = $true
			}
			#endregion
			#region Option  7) Prep Exchange Transport - RTM and CU1
			7 {
				#      Prep Exchange Transport - RTM and CU1
				Install-windowsfeature ADLDS
				Install-DotNET452
			}
			#endregion
			#region Option  8) Prep Exchange Transport - CU2+
			8 {
				#      Prep Exchange Transport - CU2+
				Install-windowsfeature ADLDS
				Install-DotNET461
			}
			#endregion
			#region Option  9) Install - One-Off - .NET 4.5.2 [MBX or Edge] - RTM or CU1
			9 {
				#      Install -One-Off - .NET 4.5.2 [MBX or Edge] - RTM or CU1
				ModuleStatus -name ServerManager
				Install-DotNET452
			}
			#endregion
			#region Option 10) Install - One-Off - .NET 4.6.1 [MBX or Edge] - CU2+
			10 {
				#      Install - One-Off - .NET 4.6.1 [MBX or Edge] - CU2+
				ModuleStatus -name ServerManager
				Install-DotNET461
				$Reboot = $true
			}
			#endregion
			#region Option 11) Install - One-Off - Windows Features [MBX]
			11 {
				#      Install -One-Off - Windows Features [MBX]
				ModuleStatus -name ServerManager
				Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
			}
			#endregion
			#region Option 12) Install - One Off - Unified Communications Managed API 4.0
			12 {
				#      Install - One Off - Unified Communications Managed API 4.0
				Install-WinUniComm4
			}
			#endregion
			#region Option 13) Prepare Schema
			13 {
				#      Prepare Schema
				Write-Verbose -Message "Insert the installation Media or Mount the Exchange ISO!!!" -verbose
				$ISOpath = Read-Host -Prompt 'Enter the installation Path <e.g. e:>'
				Write-Verbose -Message "start Prepare Schema" -verbose
				cd $ISOpath
				.\Setup.exe /PrepareSchema /IAcceptExchangeServerLicenseTerms
				Write-Host "done" -ForegroundColor green
			}
			#endregion
			#region Option 14) Prepare Active Directory and Domains
			14 {
				#      Prepare Active Directory and Domains
				Write-Verbose -Message "Insert the installation Media or Mount the Exchange ISO!!!" -verbose
				$ISOpath = Read-Host -Prompt 'Enter the installation Path <e.g. e:>'
				$domainorg = Read-Host -Prompt 'Set the name of your Exchange Organisation <e.g. Contoso>'
				$domainset = Read-Host -Prompt 'Set the name of your Domain <e.g. contoso.com>'
				Write-Verbose -Message "start Prepare Active Directory" -verbose
				cd $ISOpath
				.\Setup.exe /PrepareAD /OrganizationName: $domainorg /IAcceptExchangeServerLicenseTerms
				Write-Host "done" -ForegroundColor green
				Write-Verbose -Message "start Prepare Active Directory domains" -verbose
				cd $ISOpath
				.\Setup.exe /PrepareDomain:$domainset /IAcceptExchangeServerLicenseTerms
			}
			#endregion
			#region Option 15) Set power plan to High Performance as per Microsoft
			15 {
				#      Set power plan to High Performance as per Microsoft
				highperformance
			}
			#endregion
			#region Option 16) Disable Power Management for NICs
			16 {
				#      Disable Power Management for NICs.            
				PowerMgmt
			}
			#endregion
			#region Option 17) Disable SSL 3.0 Support
			17 {
				#      Disable SSL 3.0 Support
				DisableSSL3
			}
			#endregion
			#region Option 18) Disable RC4 Support
			18 {
				#      Disable RC4 Support       
				DisableRC4
			}
			#endregion
			#region Option 30) INSTALL EXCHANGE SERVER
			30 {
				#      INSTALL EXCHANGE SERVER
				Write-Verbose -Message "Insert the installation Media or Mount the Exchange ISO!!!" -verbose
				$ISOpath = Read-Host -Prompt 'Enter the installation Path <f.e. "e:">'
				Write-Verbose -Message "start Exchange Setup" -verbose
				cd $ISOpath
				.\setup /Mode:Install /Role:Mailbox /IAcceptExchangeServerLicenseTerms
			}
			#endregion
			#region Option 40) Add Windows Defender Exclusions
			40 {
				#   Add Windows Defender Exclusions
				ConfigurePageFile
			}
			#endregion
			#region Option 41) Show Exchange URIs
			41 {
				#   Show Exchange URIs
				ShowEXCURI
				"waiting 10 seconds..."
				sleep -Seconds 10
			}
			#endregion
			#region Option 42) Configure Exchange URLs
			42 {
				#   Configure Exchange URLs
					$server = (Get-ExchangeServer).fqdn
					$InternalURL = Read-Host "Please enter the internal URL. (Mandatory)"
					$ExternalURL = Read-Host "Please enter the external URL. (Mandatory)"
					$AutodiscoverSCP = Read-Host "Please enter the Autodiscover SCP URL. (Optional)"
					$SSLInt = Read-Host "SSL for internal Outlook Anywhere? [Y/N]"
					$SSLExt = Read-Host "SSL for external Outlook Anywhere? [Y/N]"
					
					if ($sslint -eq "y")
					{
						$InternalSSL = $true
					}
					Else
					{
						$InternalSSL = $false
					}
					if ($SSLExt -eq "y")
					{
						$ExternalSSL = $true
					}
					Else
					{
						$ExternalSSL = $false
					}
					
					ConfigureEXCURL($server, $InternalURL, $ExternalURL, $AutodiscoverSCP, $InternalSSL, $ExternalSSL)
				
			}
			#endregion
			#region Option 43) Disable UAC
			43 {
				#   Disable UAC
				New-ItemProperty -Path HKLM:Software\Microsoft\Windows\CurrentVersion\policies\system -Name EnableLUA -PropertyType DWord -Value 0 -Force
			}
			#endregion
			#region Option 44) Disable Windows Firewall
			44 {
				#   Disable Windows Firewall
				Write-Host 'Disabeling Windows Firewall...' -ForegroundColor Yellow
				Set-NetFirewallProfile -Profile Domain, Public, Private -Enabled False
				Write-Host 'Done' -ForegroundColor Green
			}
			#endregion
			#region Option 45) Create receive connector
			45 {
				#   Create receive connector
				$rcServer = Read-Host 'Enter Exchange server hostname for which to create the receive connector <e.g. Exc-srv001>'
				$rc1 = Read-Host 'Set Name of the receive connector <e.g. "Inbound SMTP Mailgateway">'
				$RemoteIPR1 = Read-Host 'Set the first Remote IP address <e.g. 000.000.000.000>'
				$RemoteIPR2 = Read-Host 'Set the second Remote IP address <e.g. 000.000.000.000>'
				Write-Host 'Setting up Receive Connector for Exchange server: $rcServer...' -ForegroundColor White
				New-ReceiveConnector -Name $rc1 -Bindings ("0.0.0.0:25") -RemoteIPRanges '$RemoteIPR1', '$RemoteIPR2' -MaxMessageSize 30MB -TransportRole FrontendTransport -Usage Custom -Server $rcServer -AuthMechanism 'TLS' -PermissionGroups 'AnonymousUsers'
				Write-Host 'Done' -ForegroundColor Green
			}
			#endregion
			#region Option 46) Create send connector
			46 {
				#   Create send connector
				$scServer = Read-Host 'Enter Exchange server hostname for which to create the send connector <e.g. Exc-srv001>'
				$sc1 = Read-Host 'Set Name of the send connector <e.g. "Outbound to Internet">'
				$sRemoteIPR1 = Read-Host 'Set the first Source Transport Server <e.g. Exc-srv001>'
				$sRemoteIPR2 = Read-Host 'Set the second Source Transport Server <e.g. Exc-srv002>'
				$sAddressSpace = Read-Host 'Set the address space: <e.g. SMTP:*.contoso.com>'
				$sSmartHost = Read-Host 'Set Smarthost <e.g. 000.000.000.000 or SM.contoso.com>'
				Write-Host 'Setting up Send Connector for Exchange server: $scServer...' -ForegroundColor White
				New-SendConnector -Name $sc1 -AddressSpaces $sAddressSpace -SourceTransportServers '$sRemoteIPR1', '$sRemoteIPR2' -FrontendProxyEnabled:$false -SmartHosts $sSmartHost
				Write-Host 'Done' -ForegroundColor Green
			}
			#endregion
			#region Option 47) Create DAG
			47 {
				#   Create DAG
				Write-Host "Starting setup to create DAG" -ForegroundColor White
				$DAGName = Read-Host 'Enter Name for DAG <e.g. DAG01>'
				$Witness = Read-Host 'Enter Hostname of Witness server <e.g. ABG-SRV01>'
				$WitnessPath = Read-Host 'Enter the local Path form Witness server where the Directory will be located: <e.g. C:\FSW\VMBDNDAGEKZ01>'
				Write-Host -Verbose 'Be sure that the Exchange permissions on the Witness server are set correctly!' -ForegroundColor Yellow -BackgroundColor Black
				$EXC01 = Read-Host 'Enter Hostname of the first Exchange server <e.g. SRV-EX01>'
				$EXC02 = Read-Host 'Enter Hostname of the second Exchange server <e.g. SRV-EX02>'
				New-DatabaseAvailabilityGroup -Name $DAGName -WitnessServer $Witness -WitnessDirectory $WitnessPath
				Add-DatabaseAvailabilityGroupServer -Identity $DAGName -MailboxServer $EXC01
				Add-DatabaseAvailabilityGroupServer -Identity $DAGName -MailboxServer $EXC02
				Get-DatabaseAvailabilityGroup $DAGName -Status
				Get-DatabaseAvailabilityGroup $DAGName -Status | fl *witness*
				Write-Host 'Done' -ForegroundColor Green
			}
			#endregion
			#region Option 48) Create Exchange Hybrid mode
			48 {
				#   Create Exchange Hybrid mode
				"This function is not yet implemented"
				
			}
			#endregion
			#region Option 49) Create Certificate request
			49 {
				#   Create Certificate request
				####################
				# Prerequisite check
				####################
				if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
				{
					Write-Host "Administrator priviliges are required. Please restart this script with elevated rights." -ForegroundColor Red
					Pause
					Throw "Administrator priviliges are required. Please restart this script with elevated rights."
				}
				
				
				#######################
				# Setting the variables
				#######################
				$UID = [guid]::NewGuid()
				$files = @{ }
				$files['settings'] = "$($env:TEMP)\$($UID)-settings.inf";
				$files['csr'] = "$($env:TEMP)\$($UID)-csr.req"
				
				
				$request = @{ }
				$request['SAN'] = @{ }
				
				Write-Host "keep it simple but significant" -ForegroundColor green
				Write-Host "Enter the Certificate informations below" -ForegroundColor cyan
				$request['CN'] = Read-Host "Common Name (e.g. company.com)"
				$request['O'] = Read-Host "Organisation (e.g. Company Ltd)"
				$request['OU'] = Read-Host "Organisational Unit (e.g. IT)"
				$request['L'] = Read-Host "City (e.g. Amsterdam)"
				$request['S'] = Read-Host "State (e.g. Noord-Holland)"
				$request['C'] = Read-Host "Country (e.g. NL)"
				
				###########################
				# Subject Alternative Names
				###########################
				$i = 0
				Do
				{
					$i++
					$request['SAN'][$i] = read-host "Subject Alternative Name $i (e.g. alt.company.com / leave empty for none)"
					if ($request['SAN'][$i] -eq "")
					{
						
					}
					
				}
				until ($request['SAN'][$i] -eq "")
				
				# Remove the last in the array (which is empty)
				$request['SAN'].Remove($request['SAN'].Count)
				
				#########################
				# Create the settings.inf
				#########################
				$settingsInf = "
                           [Version] 
                           Signature=`"`$Windows NT`$ 
                           [NewRequest] 
                           KeyLength =  2048
                           Exportable = TRUE 
                           MachineKeySet = TRUE 
                           SMIME = FALSE
                           RequestType =  PKCS10 
                           ProviderName = `"Microsoft RSA SChannel Cryptographic Provider`" 
                           ProviderType =  12
                           HashAlgorithm = sha256
                           ;Variables
                           Subject = `"CN={{CN}},OU={{OU}},O={{O}},L={{L}},S={{S}},C={{C}}`"
                           [Extensions]
                           {{SAN}}


                           ;Certreq info
                           ;http://technet.microsoft.com/en-us/library/dn296456.aspx
                           ;CSR Decoder
                           ;https://certlogik.com/decoder/
                           ;https://ssltools.websecurity.symantec.com/checker/views/csrCheck.jsp
                           "
				
				$request['SAN_string'] = & {
					if ($request['SAN'].Count -gt 0)
					{
						$san = "2.5.29.17 = `"{text}`"
"
						Foreach ($sanItem In $request['SAN'].Values)
						{
							$san += "_continue_ = `"dns=" + $sanItem + "&`"
"
						}
						return $san
					}
				}
				
				$settingsInf = $settingsInf.Replace("{{CN}}", $request['CN']).Replace("{{O}}", $request['O']).Replace("{{OU}}", $request['OU']).Replace("{{L}}", $request['L']).Replace("{{S}}", $request['S']).Replace("{{C}}", $request['C']).Replace("{{SAN}}", $request['SAN_string'])
				
				# Save settings to file in temp
				$settingsInf > $files['settings']
				
				# Done, we can start with the CSR
				Clear-Host
				
				#################################
				# CSR TIME
				#################################
				
				# Display summary
				Write-Host "Certificate information
                           Common name: $($request['CN'])
                           Organisation: $($request['O'])
                           Organisational unit: $($request['OU'])
                           City: $($request['L'])
                           State: $($request['S'])
                           Country: $($request['C'])

                           Subject alternative name(s): $($request['SAN'].Values -join ", ")

                           Signature algorithm: SHA256
                           Key algorithm: RSA
                           Key size: 2048

" -ForegroundColor Yellow
				
				certreq -new $files['settings'] $files['csr'] > $null
				
				# Output the CSR
				$CSR = Get-Content $files['csr']
				Write-Output $CSR
				Write-Host "
"
				
				# Set the Clipboard (Optional)
				Write-Host "Copy CSR to clipboard? (y|n): " -ForegroundColor Yellow -NoNewline
				if ((Read-Host) -ieq "y")
				{
					$csr | clip
					Write-Host "Check your ctrl+v
"
				}
				
				
				########################
				# Remove temporary files
				########################
				$files.Values | ForEach-Object {
					Remove-Item $_ -ErrorAction SilentlyContinue
				}
			}
			#endregion
			#region Option 50) Set mailaddress policies
			50 {
				#   set mailaddress policies
				Write-Host "Enter your Domains to create the Mail address Policies" -ForegroundColor White
				$Name1 = Read-Host "Enter the Name you wanna use for the internal Policy e.g. fabrikam-local"
				$Dom1 = Read-Host "Enter your internal Exchange Domain e.g. fabrikam.local"
				$Name2 = Read-Host "Enter the Name you wanna use for primary external Domain Policy e.g. fabrikam-extern"
				$Dom2 = Read-Host "Enter your primary external Exchange Domain e.g. fabrikam.com"
				$LocPat2 = Read-Host "Enter the Member Group OU e.g. OU=FABRIKAM,OU=CUSTOMERS,DC=fabrikam,DC=local"
				$Name3 = Read-Host "Enter the Name you wanna use for primary Accepted Domain Policy e.g. contoso-extern"
				$Dom3 = Read-Host "Enter your first Accepted Exchange Domain e.g. contoso.com"
				$LocPat3 = Read-Host "Enter the Member Group OU e.g. OU=CONTOSO,OU=CUSTOMERS,DC=fabrikam,DC=local"
				$Name4 = Read-Host "Enter the Name you wanna use for secondary Accepted Domain Policy e.g. abstergo-extern"
				$Dom4 = Read-Host "Enter your second Accepted Exchange Domain e.g. abstergo.ch"
				$LocPat4 = Read-Host "Enter the Member Group OU e.g. OU=ABSTERGO,OU=CUSTOMERS,DC=fabrikam,DC=local"
				# Create Mailaddress Policy for Resources
				Write-Host "Creating the 1st Mailaddress Policy $Name1 for Resources..." -ForegroundColor cyan
				New-EmailAddressPolicy -Name $Name1 -EnabledPrimarySMTPAddressTemplate 'SMTP:alias@$Dom1' -IncludedRecipients 'Resources' -Priority 1
				Write-Host "Done!" -ForegroundColor green
				
				# Create primary Mailaddress Policy
				Write-Host "Creating the 2nd Mailaddress Policy $Name2 for the Domain $Dom2 ..." -ForegroundColor cyan
				New-EmailAddressPolicy -Name $Name2 -EnabledPrimarySMTPAddressTemplate 'SMTP:%g.%i.%s@$Dom2' -RecipientFilter { ((MemberOfGroup -eq $LocPat2) -and (RecipientType -eq 'UserMailbox')) } -Priority 2
				Set-EmailAddressPolicy $Name2 -EnabledEmailAddressTemplates SMTP:%g.%i.%s@$Dom2, smtp:%g.%i.%s@$Dom1, smtp:%1g.%s@$Dom1, smtp:alias@$Dom1
				Write-Host "Done!" -ForegroundColor green
				
				# Create first Accepted Domain Mailaddress Policy
				Write-Host "Creating the 3rd Mailaddress Policy $Name3 for the Accepted Domain $Dom3 ..." -ForegroundColor cyan
				New-EmailAddressPolicy -Name $Name3 -EnabledPrimarySMTPAddressTemplate 'SMTP:%g.%i.%s@$Dom3' -RecipientFilter { ((MemberOfGroup -eq $LocPat3) -and (RecipientType -eq 'UserMailbox')) } -Priority 3
				Set-EmailAddressPolicy $Name3 -EnabledEmailAddressTemplates SMTP:%g.%i.%s@$Dom3, smtp:%g.%i.%s@$Dom1, smtp:%1g.%s@$Dom1, smtp:alias@$Dom1
				Write-Host "Done!" -ForegroundColor green
				
				# Create first Accepted Domain Mailaddress Policy
				Write-Host "Creating the 4th Mailaddress Policy $Name4 for the Accepted Domain $Dom4 ..." -ForegroundColor cyan
				New-EmailAddressPolicy -Name $Name4 -EnabledPrimarySMTPAddressTemplate 'SMTP:%g.%i.%s@certum.ch' -RecipientFilter { ((MemberOfGroup -eq $LocPat4) -and (RecipientType -eq 'UserMailbox')) } -Priority 4
				Set-EmailAddressPolicy $Name4 -EnabledEmailAddressTemplates SMTP:%g.%i.%s@$Dom4, smtp:%g.%i.%s@$Dom1, smtp:%1g.%s@$Dom1, smtp:alias@$Dom1
				Write-Host "Done!" -ForegroundColor green
				
				# Enable all Policies
				Write-Host "Enable all Address Policies..." -ForegroundColor cyan
				Get-EmailAddressPolicy $Name1 | Update-EmailAddressPolicy
				Get-EmailAddressPolicy $Name2 | Update-EmailAddressPolicy
				Get-EmailAddressPolicy $Name3 | Update-EmailAddressPolicy
				Get-EmailAddressPolicy $Name4 | Update-EmailAddressPolicy
				Write-Host "Done!" -ForegroundColor green
				
				# Information
				Write-Host "All Mailaddress Policies are created! See the Summary below..." -ForegroundColor magenta -Verbose
				Get-EmailAddressPolicy
			}
			#endregion
			#region Option 51) Enable UM for all Mailboxes
			51 {
				#   Enable UM for all Mailboxes
				#region log file
				$date = (Get-Date -Format yyyyMMdd_HHmm) #create time stamp
				$log = "$PSScriptRoot\$date-EnableUM.Log" #define path and name, incl. time stamp for log file
				#endregion
				
				Write-Host '--- keep it simple, but significant ---' -ForegroundColor magenta
				
				#region environment selection, modules and credentials
				#show options for environment selection
				$ExcOpt = Read-Host "Choose environment to connect to. 
                           [1] O365 
                           [2] On-Premises
                           Your option"
				
				#get credentials according to the selected environment
				switch ($ExcOpt)
				{
					1  {
						#If 1 is selected
						
						try
						{
							# Check if AzureAD module is available and import it. 
							if (!(Get-Module AzureAD) -or !(Get-Module AzureADPreview))
							{
								Import-Module AzureAD -ErrorAction Stop
								$AADModule = 'AAD'
								(Get-Date -Format G) + " " + "Azure AD module loaded" | Tee-Object -FilePath $log -Append
								
							}
							
						}
						
						catch
						{
							
							cls
							try
							{
								#Try to load MSonline module, if Azure AD module is not available
								Import-Module MSOnline -ErrorAction Stop
								$AADModule = 'MSOnline'
								(Get-Date -Format G) + " " + "MSOnline module loaded" | Tee-Object -FilePath $log -Append
							}
							catch
							{
								#If no module is available, show option to open download page for MSonline module
								Write-Host "For O365 environments you first need to install MSOnline, AzureAD, or AzureADPreview module!" -ForegroundColor Red
								Write-Host "Please install one of the modules and restart the script." -ForegroundColor Cyan
								""
								
								$red = Read-Host "Do you want to be redirected to the MS download page for the MSOnline module? [Y] Yes, [N] No. Default is No."
								switch ($red)
								{
									Y { [system.Diagnostics.Process]::Start('http://connect.microsoft.com/site1164/Downloads/DownloadDetails.aspx?DownloadID=59185') }
									N { "Script will end now." }
									default { "Script will end now." }
								}
								return
							}
							
							
							
						}
						#Ask for O365 credentials
						"O365 selected"; $O365Creds = Get-Credential -Message 'Enter your O365 credentials'
						
					}
					2  {
						#If 2 is selected, ask for On-Prem credentials
						"On-Prem selected"; $OnPremCreds = Get-Credential -Message 'Enter your Exchange On-Prem credentials'
					}
					default { Write-Host "Please enter 1, or 2" -ForegroundColor Red; return }
				}
				#endregion
				
				#region select recipient type
				
				$patwrong = $false
				$YN = $null
				do
				{
					do
					{
						#show options for recipient type detail selection
						$RecType = Read-Host "Please select the recipient type(s) you want to include. Separate multiple values by comma (1,2,...).

        [1] User Mailbox 
        [2] Shared Mailbox
        [3] Room Mailbox
        [4] Team Mailbox
        [5] Group Mailbox
        [C] Cancel

        Your selection"
						""
						
						if ($RecType -eq 'c')
						{
							"Exiting..."
							return
						}
						#Verify entered value
						$pattern = '^(?!.*?([1-5]).*?\1)[1-5](?:,[1-5])*$'
						
						if ($RecType.Length -gt 9 -or $RecType -notmatch $pattern)
						{
							'Incorrect format!'
							sleep -Seconds 1
							$patwrong = $true
							
							#return
						}
						else
						{
							$patwrong = $false
						}
					}
					until ($patwrong -eq $false)
					#Create string for get-mailbox -recipienttypedetails parameter according to user selection
					$RecType = $RecType.Replace('1', 'UserMailbox').Replace('2', 'SharedMailbox').Replace('3', 'RoomMailbox').Replace('4', 'TeamMailbox').Replace('5', 'GroupMailbox')
					#Ask if selected types are correct
					"Following recipient type(s) will be included:"
					""
					$($RecType -split ',')
					""
					
					$YN = read-host "Correct? [Y/N]"
				}
				until ($yn -eq 'y')
				
				#endregion
				
				#region set extension length
				cls
				#Ask for length of extension number
				$ExtLen = Read-Host "Please enter the length of the extension number in your environment for UM"
				#Check if a valid digit was entered
				if ($ExtLen -notmatch "\d" -or $ExtLen -eq 0)
				{
					Write-Host "Unsupported format. Only digits greater then 0 are supported." -ForegroundColor Red
					return
				}
				#endregion
				
				#region O365 | Connect
				if ($ExcOpt -eq 1)
				{
					#Connect to Exchange Online remotely
					try
					{
						$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop
						Write-Host "Connecting to Exchange Online..." -ForegroundColor Green
						Import-PSSession $Session -ErrorAction Stop | Out-Null
						(Get-Date -Format G) + " " + "Exchange Online connected" | Tee-Object -FilePath $log -Append
					}
					catch
					{
						(Get-Date -Format G) + " " + "ERROR: " + $_.exception.message | Tee-Object -FilePath $log -Append
						return
					}
					#Connect to AzureAD
					try
					{
						Write-Host "Connecting to Azure AD..." -ForegroundColor Green
						
						if ($AADModule -eq 'AAD')
						{
							#Use AzureAD module
							$aad = Connect-AzureAD -Credential $O365Creds -ErrorAction Stop
							
						}
						else
						{
							#Use MSonline module
							connect-MsolService -credential $O365Creds -ErrorAction Stop
							
						}
						
						(Get-Date -Format G) + " " + "Azure AD $($aad.TenantDomain) connected" | Tee-Object -FilePath $log -Append
						
					}
					catch
					{
						(Get-Date -Format G) + " " + "ERROR: " + $_.exception.message | Tee-Object -FilePath $log -Append
						return
					}
					
				}
				#endregion
				
				#region On-Premises | connect to Exchange
				if ($ExcOpt -eq 2)
				{
					#Ask for Exchange server name
					$Exchange = Read-Host "Enter FQDN, or short name of on-premises Exchange server. E.g. ""EXCSRV01.contoso.com, or EXCSRV01"
					
					try
					{
						#Remote connect to Exchange On-Prem 
						$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchange/PowerShell/ -Authentication Kerberos -Credential $OnPremCreds
						Import-PSSession $Session -ErrorAction Stop | Out-Null
						(Get-Date -Format G) + " " + "Exchange connected" | Tee-Object -FilePath $log -Append
					}
					catch
					{
						(Get-Date -Format G) + " " + "ERROR: " + $_.exception.message | Tee-Object -FilePath $log -Append
						Return
					}
					
				}
				#endregion
				
				#region Select UMPolicy
				cls
				#Get all UM policies
				$UMPolicies = Get-UMMailboxPolicy
				
				#Count and list UM policies 
				$count = 0
				foreach ($policy in $UMPolicies)
				{
					$count++
					Write-Host "[$count] - $($policy.Name)" -ForegroundColor Cyan
					
				}
				#Ask for UM policy to choose. (Enter number)
				[INT]$Idx = Read-Host "Enter number of UM policy to choose"
				
				#Check if entered number is valid
				if ($Idx -eq 0 -or $Idx -gt $count)
				{
					#If entered number is not valid, end script
					cls
					Write-Host "Please select a number between 1 and $count. Script ends now." -ForegroundColor Red
					return
				}
				else
				{
					#Select UM policy on base of entered number
					$UMPolicy = $UMPolicies[$idx - 1].Name
					cls
					Write-Host "You have selected the following policy: $UMPolicy" -BackgroundColor Blue
				}
				#endregion
				
				#region Enable UM users
				Write-Host "Fetching mailboxes..." -ForegroundColor Green
				#Get all mailboxes where UM is not enabled
				$mbxs = get-mailbox -RecipientTypeDetails $RecType -ResultSize unlimited | where { $_.UMEnabled -eq $false }
				$mcount = 0
				$successcount = 0
				$errorcount = 0
				
				#Go through all found mail boxes
				foreach ($mbx in $mbxs)
				{
					#Create progress bar
					$mcount++
					$percent = "{0:N1}" -f ($mcount / $mbxs.count * 100)
					Write-Progress -Activity "Enabling UM" -status "Enabling Service for $($mbx.PrimarySMTPAddress)" -percentComplete $percent -CurrentOperation "Percent completed: $percent% (no. $mcount) of $($mbxs.count) mailboxes"
					
					#Get phone number of user
					try
					{
						switch ($ExcOpt)
						{
							1 {
								#If O365 selected, get phone number from Azure AD 
								if ($AADModule -eq 'AAD')
								{
									#Use AzureAD module
									$aadUser = get-azureADUser -SearchString $mbx.UserPrincipalName -erroraction Stop
									$phone = $aadUser.TelephoneNumber
									#Throw an error if no phone number was found for the user
									if ($phone -eq "" -or $phone -eq $null)
									{
										throw "$($mbx.UserPrincipalName) - No phone number found"
										return
									}
									if ($aadUser.AssignedPlans.service -notcontains 'MicrosoftCommunicationsOnline')
									{
										throw "Error: $($mbx.userprincipalname) has no S4B Online (Plan 2) plan assigned."
										return
									}
									if ($aadUser.AssignedPlans.service -notcontains 'exchange')
									{
										throw "Error: $($mbx.userprincipalname) has no Exchange Online (E1, or E2) plan assigned."
										return
									}
								}
								else
								{
									#Use MSOnline module
									$aadUser = get-MsolUser -SearchString $mbx.UserPrincipalName -erroraction Stop
									$phone = $aadUser.PhoneNumber
									#Throw an error if no phone number was found for the user
									if ($phone -eq "" -or $phone -eq $null)
									{
										throw "$($mbx.UserPrincipalName) - No phone number found"
										return
									}
								}
							}
							2 {
								#If On-Prem is selected, use ADSI searcher to get the phone number
								$b = [adsisearcher]::new("userprincipalname=$($mbx.UserPrincipalName)")
								$result = $b.FindOne()
								$phone = $result.Properties.telephonenumber
								#Throw an error if no phone number was found for the user
								if ($phone -eq "" -or $phone -eq $null)
								{
									throw "$($mbx.UserPrincipalName) - No phone number found"
									return
								}
								
							}
						}
						
						
						#LineURI string modifiy for extension number (get only the last digits that were defined in the beginning)
						$str = $phone.TrimStart("tel:+").replace(" ", "") #Trim all spaces
						$length = $str.Length #Get length of the string
						$URI = $str.Substring(($length - $ExtLen)) #Select only substring starting from string length minus defined length
						
						#Create extension mapping (maybe used for future versions)
						$ExtensionMap = @{
							User = $mbx.PrimarySMTPAddress
							Extension = $URI
						}
						#Enable UM for the mailbox
						Enable-UMMailbox -Identity $mbx.PrimarySMTPAddress -UMMailboxPolicy $UMPolicy -SIPResourceIdentifier $mbx.PrimarySMTPAddress`
										 -Extensions $ExtensionMap.Extension -PinExpired $false -ErrorAction Stop #-WhatIf
						
						#Log
						$datetime = (Get-Date -Format G)
						"$datetime SUCCESS: $($mbx.UserPrincipalName) has been enabled for UM" | Tee-Object $log -Append
						
						#Count successfully enabled mailboxes
						$successcount++
						
					}
					catch
					{
						#Log error
						$datetime = (Get-Date -Format G)
						"$datetime ERROR: $($mbx.UserPrincipalName)  $($_.Exception.Message)" | Tee-Object $log -Append
						#Count errors
						$errorcount++
					}
					
					
				}
				#End progress bar
				Write-Progress -Activity "Enabling UM" -Completed
				
				#endregion
				
				#region show summary
				#Number of successes
				Write-Host "$successcount of $($mbxs.count) mailboxes have been successfully enabled for UM! " -ForegroundColor Green
				#If errors occurred show number of errors
				if ($errorcount -gt 0)
				{
					Write-Host "Number of errors during execution: $errorcount. Please check the log ""$log"" for details." -ForegroundColor Green
				}
				"Press any key to exit"
				cmd /c pause | Out-Null
				#endregion
			}
			#endregion
			#region Option 52) Remove old EAS devices
			52 {
				#   Remove old EAS devices
				{
					$age = $null
					
					#user input age
					$pattern = "\d+"
					Do
					{
						$Age = Read-Host "Please specify max number of days. Older entries will be removed (leave empty to cancel)"
						cls
						if ($age -notmatch $pattern -and $age -ne "")
						{
							write-host "Please enter a valid number in number format, or ""C"" to cancel!" -ForegroundColor Yellow
							sleep -Seconds 2
							cls
						}
						
					}
					Until ($age -eq "" -or $age -match $pattern -or $age -eq "c")
					if ($age -eq "")
					{
						"Cancelled"
						return
					}
					
					
					# Variables
					$now = Get-Date #Used for timestamps
					$date = $now.ToShortDateString() #Short date format for email message subject
					
					$report = @()
					
					$stats = @("DeviceID",
						"DeviceAccessState",
						"DeviceAccessStateReason",
						"DeviceModel"
						"DeviceType",
						"DeviceFriendlyName",
						"DeviceOS",
						"LastSyncAttemptTime",
						"LastSuccessSync"
					)
					
					$reportemailsubject = "Exchange ActiveSync Device Report - $date"
					$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
					$reportfile = "$myDir\ExchangeActiveSyncDevice-ToDelete.csv"
					
					
					
					
					
					# Initialize
					#Add Exchange 2010/2013/2016 snapin if not already loaded in the PowerShell session
					if (!(Get-PSSnapin | where { $_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010" }))
					{
						try
						{
							Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
						}
						catch
						{
							#Snapin was not loaded
							Write-Warning $_.Exception.Message
							return
						}
						. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
						Connect-ExchangeServer -auto -AllowClobber
					}
					
					
					
					# Script
					Write-Host "keep it simple but significant" -ForegroundColor magenta
					Start-Sleep -s 2
					Write-Host "Fetching List of Mailboxes with EAS Device partnerships" -ForegroundColor cyan
					Start-Sleep -s 5
					Write-Host "Don't worry, this can take a while..." -ForegroundColor cyan
					
					$MailboxesWithEASDevices = @(Get-CASMailbox -Resultsize Unlimited | Where { $_.HasActiveSyncDevicePartnership })
					
					Write-Host "$($MailboxesWithEASDevices.count) mailboxes with EAS device partnerships"
					
					$i = 0
					
					Foreach ($Mailbox in $MailboxesWithEASDevices)
					{
						
						$EASDeviceStats = @(Get-ActiveSyncDeviceStatistics -Mailbox $Mailbox.Identity -WarningAction SilentlyContinue)
						
						Write-Host "$($Mailbox.Identity) has $($EASDeviceStats.Count) device(s)"
						
						$MailboxInfo = Get-Mailbox $Mailbox.Identity | Select DisplayName, PrimarySMTPAddress, OrganizationalUnit
						
						Foreach ($EASDevice in $EASDeviceStats)
						{
							Write-Host -ForegroundColor Green "Processing $($EASDevice.DeviceID)"
							
							$lastsyncattempt = ($EASDevice.LastSyncAttemptTime)
							
							if ($lastsyncattempt -eq $null)
							{
								$syncAge = "Never"
							}
							else
							{
								$syncAge = ($now - $lastsyncattempt).Days
							}
							
							#Add to report if last sync attempt greater than Age specified
							if ($syncAge -ge $Age -or $syncAge -eq "Never" -and $EASDevice.DeviceID -ne 0)
							{
								Write-Host -ForegroundColor Yellow "$($EASDevice.DeviceID) sync age of $syncAge days is greater than $age, adding to report"
								
								$reportObj = New-Object PSObject
								$reportObj | Add-Member NoteProperty -Name "Display Name" -Value $MailboxInfo.DisplayName
								$reportObj | Add-Member NoteProperty -Name "Organizational Unit" -Value $MailboxInfo.OrganizationalUnit
								$reportObj | Add-Member NoteProperty -Name "Email Address" -Value $MailboxInfo.PrimarySMTPAddress
								$reportObj | Add-Member NoteProperty -Name "Sync Age (Days)" -Value $syncAge
								$reportObj | Add-Member NoteProperty -Name "GUID" -Value $EASDevice.GUID
								
								Foreach ($stat in $stats)
								{
									$reportObj | Add-Member NoteProperty -Name $stat -Value $EASDevice.$stat
								}
								
								$report += $reportObj
							}
						}
						$i++
						Write-Progress -activity "Gethering EAS devices . . ." -status "Collected: $i of $($MailboxesWithEASDevices.Count)" -percentComplete (($i / $MailboxesWithEASDevices.Count) * 100)
					}
					Write-Progress -activity "Gethering EAS devices . . ." -Completed
					
					Write-Host -ForegroundColor White "Saving report to $reportfile"
					$report | Export-Csv -NoTypeInformation $reportfile -Encoding UTF8
					
					ii $reportfile #Open the CSV. File 
					Write-Host "!!! with great power comes great responsibility !!!" -ForegroundColor magenta
					Write-Host "Check the CSV File before you continue! To continue push ENTER" -ForegroundColor Gray -NoNewline
					$dummy = Read-Host
					
					$ReportToDelete = Import-csv $reportfile
					###
					
					$counter = 0
					$sum = $ReportToDelete.count
					foreach ($i in $ReportToDelete)
					{
						try
						{
							write-host $i."Display Name" $i."LastSuccessSync" $i."DeviceFriendlyName"
							Remove-MobileDevice -Identity $i."GUID" -Confirm:$false -erroraction Stop #Remove the selected MobileDevices (by GUID)
							Write-Host "Device removed" -ForegroundColor Green
							(get-date -Format g) + " Success: Removed device: " + $i."Display Name" + $i."DeviceFriendlyName" | Out-File $PSScriptRoot\Successlog.log -Append
						}
						catch
						{
							
							(get-date -Format g) + " Error: " + $i."Display Name" + $i."DeviceFriendlyName" + " " + $_.exception.message | Out-File $PSScriptRoot\errorlog.log -Append
							Write-Host "Error while removing device" -ForegroundColor Red
							$_.exception.message
						}
						$counter++
						Write-Progress -activity "Removing EAS devices . . ." -status "Processed: $counter of $($sum)" -percentComplete (($counter / $sum) * 100)
					}
					Write-Progress -Activity "Removing EAS devices . . ." -Completed
					
					Write-Host "Active sync Devices older then $Age Days are successfully removed for the Exchange Organization" -ForegroundColor green
				}
			}#endregion
			#region Option 53) Deploy Microsoft Teams Desktop Client
			53 {
				"This function is not yet implemented"
				#      Deploy Microsoft Teams Desktop Client
                           <#
.SYNOPSIS
Install-MicrosoftTeams.ps1 - Microsoft Teams Desktop Client Deployment Script

.DESCRIPTION 
This PowerShell script will silently install the Microsoft Teams desktop client.

The Teams client installer can be downloaded from Microsoft:
https://teams.microsoft.com/downloads

.PARAMETER SourcePath
Specifies the source path for the Microsoft Teams installer.


.EXAMPLE
.\Install-MicrosoftTeams.ps1 -Source \\mgmt\Installs\MicrosoftTeams

Installs the Microsoft Teams client from the Installs share on the server MGMT.

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:   http://paulcunningham.me
* Twitter:   https://twitter.com/paulcunningham
* LinkedIn:  http://au.linkedin.com/in/cunninghamp/
* Github:    https://github.com/cunninghamp

For more Office 365 tips, tricks and news
check out Practical 365.

* Website:   http://practical365.com
* Twitter:   http://twitter.com/practical365

Change Log
V1.00, 15/03/2017 - Initial version
#>
				
				
				<##requires -version 4
				{
					[CmdletBinding()]
					param (
						
						[Parameter(Mandatory = $true)]
						[string]$SourcePath
						
					)
					
					
					function DoInstall
					{
						
						$Installer = "$($SourcePath)\Teams_windows_x64.exe"
						
						If (!(Test-Path $Installer))
						{
							throw "Unable to locate Microsoft Teams client installer at $($installer)"
						}
						
						Write-Host "Attempting to install Microsoft Teams client"
						
						try
						{
							$process = Start-Process -FilePath "$Installer" -ArgumentList "-s" -Wait -PassThru -ErrorAction STOP
							
							if ($process.ExitCode -eq 0)
							{
								Write-Host -ForegroundColor Green "Microsoft Teams setup started without error."
							}
							else
							{
								Write-Warning "Installer exit code  $($process.ExitCode)."
							}
						}
						catch
						{
							Write-Warning $_.Exception.Message
						}
						
					}
					
					#Check if Office is already installed, as indicated by presence of registry key
					$installpath = "$($env:LOCALAPPDATA)\Microsoft\Teams"
					
					if (-not (Test-Path "$($installpath)\Update.exe"))
					{
						DoInstall
					}
					else
					{
						if (Test-Path "$($installpath)\.dead")
						{
							Write-Host "Teams was previously installed but has been uninstalled. Will reinstall."
							DoInstall
						}
					}
					
					
					
				}
				
				
				$Reboot = $true#>
			}
			#endregion
			#region Option 54) Order certificate >>GO DADDY<<
			54 {
				"This function is not yet implemented"
				#      Order certificate >>GO DADDY<<
				#Order certificate
				#empty
				<#
				#Variables
				$cersrv = Read-Host 'Enter the server name where you wanna import the certificate <e.g. EXCsrv01>'
				$cerpath = Read-Host 'Enter the the path, where your certificate is located <e.g. \\FileServer01\Data\>'
				$cercertname = Read-Host 'Enter the Name of the certificate <e.g. 'Exported Fabrikam Cert.pfx'>'
				$cerPW = Read-Host 'Enter the Password of the .PFX file PLEASE NOTE, YOU ENTER IT AS PLAIN TEXT! the password will be converted to a Secure String automaticaly!'
				
				#Script import certificate
				Import-ExchangeCertificate -Server $cersrv -FileName "$cerpath", "$cercertname" -Password (ConvertTo-SecureString -String $cerPW -AsPlainText -Force)
				
				$Reboot = $false#>
			}
			#endregion
			#region Option 55) Order certificate >>DIGICERT<<
			55 {
				"This function is not yet implemented"
				#      Order certificate >>DIGICERT<<
				#Order certificate
				#empty
				<#
				#Variables
				$cersrv = Read-Host 'Enter the server name where you wanna import the certificate <e.g. EXCsrv01>'
				$cerpath = Read-Host 'Enter the the path, where your certificate is located <e.g. \\FileServer01\Data\>'
				$cercertname = Read-Host 'Enter the Name of the certificate <e.g. 'Exported Fabrikam Cert.pfx'>'
				$cerPW = Read-Host 'Enter the Password of the .PFX file PLEASE NOTE, YOU ENTER IT AS PLAIN TEXT! the password will be converted to a Secure String automaticaly!'
				
				#Script import certificate
				Import-ExchangeCertificate -Server $cersrv -FileName "$cerpath", "$cercertname" -Password (ConvertTo-SecureString -String $cerPW -AsPlainText -Force)
				
				$Reboot = $false#>
			}
			#endregion
			#region Option 60) Generate Health Report for an Exchange Server 2016/2013/2010 Environment
			60 {
				generateHealthReport
			}
			#endregion
			#region Option 61) Generate Exchange Environment Reports
			61 {
				#      Generate Exchange Environment Reports
                           <#
    .SYNOPSIS
    Creates a HTML Report describing the Exchange environment 
   
       Steve Goodman
       
       THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
       RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
       
       Version 1.6.2 January 2017
       
    .DESCRIPTION
       
    This script creates a HTML report showing the following information about an Exchange 
    2016, 2013, 2010 and to a lesser extent, 2007 and 2003, environment. 
    
    The following is shown:
       
       * Report Generation Time
       * Total Servers per Exchange Version (2003 > 2010 or 2007 > 2016)
       * Total Mailboxes per Exchange Version, Office 365 and Organisation
       * Total Roles in the environment
             
       Then, per site:
       * Total Mailboxes per site
    * Internal, External and CAS Array Hostnames
       * Exchange Servers with:
             o Exchange Server Version
             o Service Pack
             o Update Rollup and rollup version
             o Roles installed on server and mailbox counts
             o OS Version and Service Pack
             
       Then, per Database availability group (Exchange 2010/2013/2016):
       * Total members per DAG
       * Member list
       * Databases, detailing:
             o Mailbox Count and Average Size
             o Archive Mailbox Count and Average Size (Only shown if DAG includes Archive Mailboxes)
             o Database Size and whitespace
             o Database and log disk free
             o Last Full Backup (Only shown if one or more DAG database has been backed up)
             o Circular Logging Enabled (Only shown if one or more DAG database has Circular Logging enabled)
             o Mailbox server hosting active copy
             o List of mailbox servers hosting copies and number of copies
             
       Finally, per Database (Non DAG DBs/Exchange 2007/Exchange 2003)
       * Databases, detailing:
             o Storage Group (if applicable) and DB name
             o Server hosting database
             o Mailbox Count and Average Size
             o Archive Mailbox Count and Average Size (Only shown if DAG includes Archive Mailboxes)
             o Database Size and whitespace
             o Database and log disk free
             o Last Full Backup (Only shown if one or more DAG database has been backed up)
             o Circular Logging Enabled (Only shown if one or more DAG database has Circular Logging enabled)
             
       This does not detail public folder infrastructure, or examine Exchange 2007/2003 CCR/SCC clusters
       (although it attempts to detect Clustered Exchange 2007/2003 servers, signified by ClusMBX).
       
       IMPORTANT NOTE: The script requires WMI and Remote Registry access to Exchange servers from the server 
       it is run from to determine OS version, Update Rollup, Exchange 2007/2003 cluster and DB size information.
       
       .PARAMETER HTMLReport
    Filename to write HTML Report to
       
       .PARAMETER SendMail
       Send Mail after completion. Set to $True to enable. If enabled, -MailFrom, -MailTo, -MailServer are mandatory
       
       .PARAMETER MailFrom
       Email address to send from. Passed directly to Send-MailMessage as -From
       
       .PARAMETER MailTo
       Email address to send to. Passed directly to Send-MailMessage as -To
       
       .PARAMETER MailServer
       SMTP Mail server to attempt to send through. Passed directly to Send-MailMessage as -SmtpServer
       
       .PARAMETER ScheduleAs
       Attempt to schedule the command just executed for 10PM nightly. Specify the username here, schtasks (under the hood) will ask for a password later.
    
       .PARAMETER ViewEntireForest
       By default, true. Set the option in Exchange 2007 or 2010 to view all Exchange servers and recipients in the forest.
   
    .PARAMETER ServerFilter
       Use a text based string to filter Exchange Servers by, e.g. NL-* -  Note the use of the wildcard (*) character to allow for multiple matches.
    
       .EXAMPLE
    Generate the HTML report 
    .\Get-ExchangeEnvironmentReport.ps1 -HTMLReport .\report.html
       
    #>
				#region user input
				$HTMLReport = Read-Host "Specifiy alternate path and name for report file. Default is ""exchangeenvironmentreport.html"""
				if ($HTMLReport = "")
				{
					$ReportFile = "exchangeenvironmentreport.html"
				}
				$SendMailYesNo = Read-Host "Send e-mail with report? [Y/N] Default is [N]"
				
				switch ($SendMailYesNo)
				{
					Y{ $SendEmail = $true }
					N{ $SendEmail = $false }
					default { "No option selected. Exiting"; Return }
				}
				if ($SendEmail)
				{
					$AlertsOnlyYN = Read-Host "Send email only if error or warning was detected?[Y/N] Default is [N]"
					switch ($AlertsOnlyYN)
					{
						Y{ $AlertsOnly = $true }
						N{ $AlertsOnly = $false }
						default { $AlertsOnly = $false }
					}
					$MailServer = Read-Host "Enter SMTP Server"
					$MailTo = Read-Host -Prompt "Enter recipients SMTP address"
					$MailFrom = Read-Host -Prompt "Enter senders SMTP address"
					
				}
				$ScheduleAsYN = Read-Host "Do you want to schedule the execution?[Y/N]"
				if ($ScheduleAsYN = "Y")
				{
					$ScheduleAs = Read-Host "Enter username in which context the scheduled will run"
					if ($ScheduleAs = "")
					{
						Write-Host "No username specified. Exiting now."
						Return
					}
				}
				$ViewEntireForestYN = Read-Host "View entire forest (all Exchange servers and recipients)[Y/N] Default is [Y]"
				switch ($ViewEntireForestYN)
				{
					Y{ $ViewEntireForest = $true }
					N{ $ViewEntireForest = $false }
					default { $AlertsOnly = $true }
				}
				$ServerFilter = Read-Host "Specifiy a filter for server names. Wildcards are allowed. Default is ""*"""
				if ($ServerFilter = "")
				{
					$ServerFilter = "*"
				}
				#endregion
				
				# Sub-Function to Get Database Information. Shorter than expected..
				function _GetDAG
				{
					param ($DAG)
					@{
						Name = $DAG.Name.ToUpper()
						MemberCount = $DAG.Servers.Count
						Members = [array]($DAG.Servers | % { $_.Name })
						Databases = @()
					}
				}
				
				
				# Sub-Function to Get Database Information
				function _GetDB
				{
					param ($Database,
						$ExchangeEnvironment,
						$Mailboxes,
						$ArchiveMailboxes,
						$E2010)
					
					# Circular Logging, Last Full Backup
					if ($Database.CircularLoggingEnabled) { $CircularLoggingEnabled = "Yes" }
					else { $CircularLoggingEnabled = "No" }
					if ($Database.LastFullBackup) { $LastFullBackup = $Database.LastFullBackup.ToString() }
					else { $LastFullBackup = "Not Available" }
					
					# Mailbox Average Sizes
					$MailboxStatistics = [array]($ExchangeEnvironment.Servers[$Database.Server.Name].MailboxStatistics | Where { $_.Database -eq $Database.Identity })
					if ($MailboxStatistics)
					{
						[long]$MailboxItemSizeB = 0
						$MailboxStatistics | %{ $MailboxItemSizeB += $_.TotalItemSizeB }
						[long]$MailboxAverageSize = $MailboxItemSizeB / $MailboxStatistics.Count
					}
					else
					{
						$MailboxAverageSize = 0
					}
					
					# Free Disk Space Percentage
					if ($ExchangeEnvironment.Servers[$Database.Server.Name].Disks)
					{
						foreach ($Disk in $ExchangeEnvironment.Servers[$Database.Server.Name].Disks)
						{
							if ($Database.EdbFilePath.PathName -like "$($Disk.Name)*")
							{
								$FreeDatabaseDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
							}
							if ($Database.ExchangeVersion.ExchangeBuild.Major -ge 14)
							{
								if ($Database.LogFolderPath.PathName -like "$($Disk.Name)*")
								{
									$FreeLogDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
								}
							}
							else
							{
								$StorageGroupDN = $Database.DistinguishedName.Replace("CN=$($Database.Name),", "")
								$Adsi = [adsi]"LDAP://$($Database.OriginatingServer)/$($StorageGroupDN)"
								if ($Adsi.msExchESEParamLogFilePath -like "$($Disk.Name)*")
								{
									$FreeLogDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
								}
							}
						}
					}
					else
					{
						$FreeLogDiskSpace = $null
						$FreeDatabaseDiskSpace = $null
					}
					
					if ($Database.ExchangeVersion.ExchangeBuild.Major -ge 14 -and $E2010)
					{
						# Exchange 2010 Database Only
						$CopyCount = [int]$Database.Servers.Count
						if ($Database.MasterServerOrAvailabilityGroup.Name -ne $Database.Server.Name)
						{
							$Copies = [array]($Database.Servers | % { $_.Name })
						}
						else
						{
							$Copies = @()
						}
						# Archive Info
						$ArchiveMailboxCount = [int]([array]($ArchiveMailboxes | Where { $_.ArchiveDatabase -eq $Database.Name })).Count
						$ArchiveStatistics = [array]($ArchiveMailboxes | Where { $_.ArchiveDatabase -eq $Database.Name } | Get-MailboxStatistics -Archive)
						if ($ArchiveStatistics)
						{
							[long]$ArchiveItemSizeB = 0
							$ArchiveStatistics | %{ $ArchiveItemSizeB += $_.TotalItemSize.Value.ToBytes() }
							[long]$ArchiveAverageSize = $ArchiveItemSizeB / $ArchiveStatistics.Count
						}
						else
						{
							$ArchiveAverageSize = 0
						}
						# DB Size / Whitespace Info
						[long]$Size = $Database.DatabaseSize.ToBytes()
						[long]$Whitespace = $Database.AvailableNewMailboxSpace.ToBytes()
						$StorageGroup = $null
						
					}
					else
					{
						$ArchiveMailboxCount = 0
						$CopyCount = 0
						$Copies = @()
						# 2003 & 2007, Use WMI (Based on code by Gary Siepser, http://bit.ly/kWWMb3)
						$Size = [long](get-wmiobject cim_datafile -computername $Database.Server.Name -filter ('name=''' + $Database.edbfilepath.pathname.replace("\", "\\") + '''')).filesize
						if (!$Size)
						{
							Write-Warning "Cannot detect database size via WMI for $($Database.Server.Name)"
							[long]$Size = 0
							[long]$Whitespace = 0
						}
						else
						{
							[long]$MailboxDeletedItemSizeB = 0
							if ($MailboxStatistics)
							{
								$MailboxStatistics | %{ $MailboxDeletedItemSizeB += $_.TotalDeletedItemSizeB }
							}
							$Whitespace = $Size - $MailboxItemSizeB - $MailboxDeletedItemSizeB
							if ($Whitespace -lt 0) { $Whitespace = 0 }
						}
						$StorageGroup = $Database.DistinguishedName.Split(",")[1].Replace("CN=", "")
					}
					
					@{
						Name = $Database.Name
						StorageGroup = $StorageGroup
						ActiveOwner = $Database.Server.Name.ToUpper()
						MailboxCount = [long]([array]($Mailboxes | Where { $_.Database -eq $Database.Identity })).Count
						MailboxAverageSize = $MailboxAverageSize
						ArchiveMailboxCount = $ArchiveMailboxCount
						ArchiveAverageSize = $ArchiveAverageSize
						CircularLoggingEnabled = $CircularLoggingEnabled
						LastFullBackup = $LastFullBackup
						Size = $Size
						Whitespace = $Whitespace
						Copies = $Copies
						CopyCount = $CopyCount
						FreeLogDiskSpace = $FreeLogDiskSpace
						FreeDatabaseDiskSpace = $FreeDatabaseDiskSpace
					}
				}
				
				
				# Sub-Function to get mailbox count per server.
				# New in 1.5.2
				function _GetExSvrMailboxCount
				{
					param ($Mailboxes,
						$ExchangeServer,
						$Databases)
					# The following *should* work, but it doesn't. Apparently, ServerName is not always returned correctly which may be the cause of
					# reports of counts being incorrect
					#([array]($Mailboxes | Where {$_.ServerName -eq $ExchangeServer.Name})).Count
					
					# ..So as a workaround, I'm going to check what databases are assigned to each server and then get the mailbox counts on a per-
					# database basis and return the resulting total. As we already have this information resident in memory it should be cheap, just
					# not as quick.
					$MailboxCount = 0
					foreach ($Database in [array]($Databases | Where { $_.Server -eq $ExchangeServer.Name }))
					{
						$MailboxCount += ([array]($Mailboxes | Where { $_.Database -eq $Database.Identity })).Count
					}
					$MailboxCount
					
				}
				
				# Sub-Function to Get Exchange Server information
				function _GetExSvr
				{
					param ($E2010,
						$ExchangeServer,
						$Mailboxes,
						$Databases,
						$Hybrids)
					
					# Set Basic Variables
					$MailboxCount = 0
					$RollupLevel = 0
					$RollupVersion = ""
					$ExtNames = @()
					$IntNames = @()
					$CASArrayName = ""
					
					# Get WMI Information
					$tWMI = Get-WmiObject Win32_OperatingSystem -ComputerName $ExchangeServer.Name -ErrorAction SilentlyContinue
					if ($tWMI)
					{
						$OSVersion = $tWMI.Caption.Replace("(R)", "").Replace("Microsoft ", "").Replace("Enterprise", "Ent").Replace("Standard", "Std").Replace(" Edition", "")
						$OSServicePack = $tWMI.CSDVersion
						$RealName = $tWMI.CSName.ToUpper()
					}
					else
					{
						Write-Warning "Cannot detect OS information via WMI for $($ExchangeServer.Name)"
						$OSVersion = "N/A"
						$OSServicePack = "N/A"
						$RealName = $ExchangeServer.Name.ToUpper()
					}
					$tWMI = Get-WmiObject -query "Select * from Win32_Volume" -ComputerName $ExchangeServer.Name -ErrorAction SilentlyContinue
					if ($tWMI)
					{
						$Disks = $tWMI | Select Name, Capacity, FreeSpace | Sort-Object -Property Name
					}
					else
					{
						Write-Warning "Cannot detect OS information via WMI for $($ExchangeServer.Name)"
						$Disks = $null
					}
					
					# Get Exchange Version
					if ($ExchangeServer.AdminDisplayVersion.Major -eq 6)
					{
						$ExchangeMajorVersion = "$($ExchangeServer.AdminDisplayVersion.Major).$($ExchangeServer.AdminDisplayVersion.Minor)"
						$ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.FilePatchLevelDescription.Replace("Service Pack ", "")
					}
					elseif ($ExchangeServer.AdminDisplayVersion.Major -eq 15 -and $ExchangeServer.AdminDisplayVersion.Minor -eq 1)
					{
						$ExchangeMajorVersion = [double]"$($ExchangeServer.AdminDisplayVersion.Major).$($ExchangeServer.AdminDisplayVersion.Minor)"
						$ExchangeSPLevel = 0
					}
					else
					{
						$ExchangeMajorVersion = $ExchangeServer.AdminDisplayVersion.Major
						$ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.Minor
					}
					# Exchange 2007+
					if ($ExchangeMajorVersion -ge 8)
					{
						# Get Roles
						$MailboxStatistics = $null
						[array]$Roles = $ExchangeServer.ServerRole.ToString().Replace(" ", "").Split(",");
						# Add Hybrid "Role" for report
						if ($Hybrids -contains $ExchangeServer.Name)
						{
							$Roles += "Hybrid"
						}
						if ($Roles -contains "Mailbox")
						{
							$MailboxCount = _GetExSvrMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
							if ($ExchangeServer.Name.ToUpper() -ne $RealName)
							{
								$Roles = [array]($Roles | Where { $_ -ne "Mailbox" })
								$Roles += "ClusteredMailbox"
							}
							# Get Mailbox Statistics the normal way, return in a consitent format
							$MailboxStatistics = Get-MailboxStatistics -Server $ExchangeServer | Select DisplayName, @{ Name = "TotalItemSizeB"; Expression = { $_.TotalItemSize.Value.ToBytes() } }, @{ Name = "TotalDeletedItemSizeB"; Expression = { $_.TotalDeletedItemSize.Value.ToBytes() } }, Database
						}
						# Get HTTPS Names (Exchange 2010 only due to time taken to retrieve data)
						if ($Roles -contains "ClientAccess" -and $E2010)
						{
							
							Get-OWAVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames += $_.ExternalURL.Host; $IntNames += $_.InternalURL.Host; }
							Get-WebServicesVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames += $_.ExternalURL.Host; $IntNames += $_.InternalURL.Host; }
							Get-OABVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames += $_.ExternalURL.Host; $IntNames += $_.InternalURL.Host; }
							Get-ActiveSyncVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames += $_.ExternalURL.Host; $IntNames += $_.InternalURL.Host; }
							if (Get-Command Get-MAPIVirtualDirectory -ErrorAction SilentlyContinue)
							{
								Get-MAPIVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames += $_.ExternalURL.Host; $IntNames += $_.InternalURL.Host; }
							}
							if (Get-Command Get-ClientAccessService -ErrorAction SilentlyContinue)
							{
								$IntNames += (Get-ClientAccessService -Identity $ExchangeServer.Name).AutoDiscoverServiceInternalURI.Host
							}
							else
							{
								$IntNames += (Get-ClientAccessServer -Identity $ExchangeServer.Name).AutoDiscoverServiceInternalURI.Host
							}
							
							if ($ExchangeMajorVersion -ge 14)
							{
								Get-ECPVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames += $_.ExternalURL.Host; $IntNames += $_.InternalURL.Host; }
							}
							$IntNames = $IntNames | Sort-Object -Unique
							$ExtNames = $ExtNames | Sort-Object -Unique
							$CASArray = Get-ClientAccessArray -Site $ExchangeServer.Site.Name
							if ($CASArray)
							{
								$CASArrayName = $CASArray.Fqdn
							}
						}
						
						# Rollup Level / Versions (Thanks to Bhargav Shukla http://bit.ly/msxGIJ)
						if ($ExchangeMajorVersion -ge 14)
						{
							$RegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches"
						}
						else
						{
							$RegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\461C2B4266EDEF444B864AD6D9E5B613\\Patches"
						}
						$RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name);
						if ($RemoteRegistry)
						{
							$RUKeys = $RemoteRegistry.OpenSubKey($RegKey).GetSubKeyNames() | ForEach { "$RegKey\\$_" }
							if ($RUKeys)
							{
								[array]($RUKeys | %{ $RemoteRegistry.OpenSubKey($_).getvalue("DisplayName") }) | %{
									if ($_ -like "Update Rollup *")
									{
										$tRU = $_.Split(" ")[2]
										if ($tRU -like "*-*") { $tRUV = $tRU.Split("-")[1]; $tRU = $tRU.Split("-")[0] }
										else { $tRUV = "" }
										if ([int]$tRU -ge [int]$RollupLevel) { $RollupLevel = $tRU; $RollupVersion = $tRUV }
									}
								}
							}
						}
						else
						{
							Write-Warning "Cannot detect Rollup Version via Remote Registry for $($ExchangeServer.Name)"
						}
						# Exchange 2013 CU or SP Level
						if ($ExchangeMajorVersion -ge 15)
						{
							$RegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Microsoft Exchange v15"
							$RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name);
							if ($RemoteRegistry)
							{
								$ExchangeSPLevel = $RemoteRegistry.OpenSubKey($RegKey).getvalue("DisplayName")
								if ($ExchangeSPLevel -like "*Service Pack*" -or $ExchangeSPLevel -like "*Cumulative Update*")
								{
									$ExchangeSPLevel = $ExchangeSPLevel.Replace("Microsoft Exchange Server 2013 ", "");
									$ExchangeSPLevel = $ExchangeSPLevel.Replace("Microsoft Exchange Server 2016 ", "");
									$ExchangeSPLevel = $ExchangeSPLevel.Replace("Service Pack ", "SP");
									$ExchangeSPLevel = $ExchangeSPLevel.Replace("Cumulative Update ", "CU");
								}
								else
								{
									$ExchangeSPLevel = 0;
								}
							}
							else
							{
								Write-Warning "Cannot detect CU/SP via Remote Registry for $($ExchangeServer.Name)"
							}
						}
						
					}
					# Exchange 2003
					if ($ExchangeMajorVersion -eq 6.5)
					{
						# Mailbox Count
						$MailboxCount = _GetExSvrMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
						# Get Role via WMI
						$tWMI = Get-WMIObject Exchange_Server -Namespace "root\microsoftexchangev2" -Computername $ExchangeServer.Name -Filter "Name='$($ExchangeServer.Name)'"
						if ($tWMI)
						{
							if ($tWMI.IsFrontEndServer) { $Roles = @("FE") }
							else { $Roles = @("BE") }
						}
						else
						{
							Write-Warning "Cannot detect Front End/Back End Server information via WMI for $($ExchangeServer.Name)"
							$Roles += "Unknown"
						}
						# Get Mailbox Statistics using WMI, return in a consistent format
						$tWMI = Get-WMIObject -class Exchange_Mailbox -Namespace ROOT\MicrosoftExchangev2 -ComputerName $ExchangeServer.Name -Filter ("ServerName='$($ExchangeServer.Name)'")
						if ($tWMI)
						{
							$MailboxStatistics = $tWMI | Select @{ Name = "DisplayName"; Expression = { $_.MailboxDisplayName } }, @{ Name = "TotalItemSizeB"; Expression = { $_.Size } }, @{ Name = "TotalDeletedItemSizeB"; Expression = { $_.DeletedMessageSizeExtended } }, @{ Name = "Database"; Expression = { ((get-mailboxdatabase -Identity "$($_.ServerName)\$($_.StorageGroupName)\$($_.StoreName)").identity) } }
						}
						else
						{
							Write-Warning "Cannot retrieve Mailbox Statistics via WMI for $($ExchangeServer.Name)"
							$MailboxStatistics = $null
						}
					}
					# Exchange 2000
					if ($ExchangeMajorVersion -eq "6.0")
					{
						# Mailbox Count
						$MailboxCount = _GetExSvrMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
						# Get Role via ADSI
						$tADSI = [ADSI]"LDAP://$($ExchangeServer.OriginatingServer)/$($ExchangeServer.DistinguishedName)"
						if ($tADSI)
						{
							if ($tADSI.ServerRole -eq 1) { $Roles = @("FE") }
							else { $Roles = @("BE") }
						}
						else
						{
							Write-Warning "Cannot detect Front End/Back End Server information via ADSI for $($ExchangeServer.Name)"
							$Roles += "Unknown"
						}
						$MailboxStatistics = $null
					}
					
					# Return Hashtable
					@{
						Name = $ExchangeServer.Name.ToUpper()
						RealName = $RealName
						ExchangeMajorVersion = $ExchangeMajorVersion
						ExchangeSPLevel = $ExchangeSPLevel
						Edition = $ExchangeServer.Edition
						Mailboxes = $MailboxCount
						OSVersion = $OSVersion;
						OSServicePack = $OSServicePack
						Roles = $Roles
						RollupLevel = $RollupLevel
						RollupVersion = $RollupVersion
						Site = $ExchangeServer.Site.Name
						MailboxStatistics = $MailboxStatistics
						Disks = $Disks
						IntNames = $IntNames
						ExtNames = $ExtNames
						CASArrayName = $CASArrayName
					}
				}
				
				# Sub Function to Get Totals by Version
				function _TotalsByVersion
				{
					param ($ExchangeEnvironment)
					$TotalMailboxesByVersion = @{ }
					if ($ExchangeEnvironment.Sites)
					{
						foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator())
						{
							foreach ($Server in $Site.Value)
							{
								if (!$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"])
								{
									$TotalMailboxesByVersion.Add("$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)", @{ ServerCount = 1; MailboxCount = $Server.Mailboxes })
								}
								else
								{
									$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].ServerCount++
									$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].MailboxCount += $Server.Mailboxes
								}
							}
						}
					}
					if ($ExchangeEnvironment.Pre2007)
					{
						foreach ($FakeSite in $ExchangeEnvironment.Pre2007.GetEnumerator())
						{
							foreach ($Server in $FakeSite.Value)
							{
								if (!$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"])
								{
									$TotalMailboxesByVersion.Add("$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)", @{ ServerCount = 1; MailboxCount = $Server.Mailboxes })
								}
								else
								{
									$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].ServerCount++
									$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].MailboxCount += $Server.Mailboxes
								}
							}
						}
					}
					$TotalMailboxesByVersion
				}
				
				# Sub Function to Get Totals by Role
				function _TotalsByRole
				{
					param ($ExchangeEnvironment)
					# Add Roles We Always Show
					$TotalServersByRole = @{
						"ClientAccess"	   = 0
						"HubTransport"	   = 0
						"UnifiedMessaging" = 0
						"Mailbox"		   = 0
						"Edge"			   = 0
					}
					if ($ExchangeEnvironment.Sites)
					{
						foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator())
						{
							foreach ($Server in $Site.Value)
							{
								foreach ($Role in $Server.Roles)
								{
									if ($TotalServersByRole[$Role] -eq $null)
									{
										$TotalServersByRole.Add($Role, 1)
									}
									else
									{
										$TotalServersByRole[$Role]++
									}
								}
							}
						}
					}
					if ($ExchangeEnvironment.Pre2007["Pre 2007 Servers"])
					{
						
						foreach ($Server in $ExchangeEnvironment.Pre2007["Pre 2007 Servers"])
						{
							
							foreach ($Role in $Server.Roles)
							{
								if ($TotalServersByRole[$Role] -eq $null)
								{
									$TotalServersByRole.Add($Role, 1)
								}
								else
								{
									$TotalServersByRole[$Role]++
								}
							}
						}
					}
					$TotalServersByRole
				}
				
				# Sub Function to return HTML Table for Sites/Pre 2007
				function _GetOverview
				{
					param ($Servers,
						$ExchangeEnvironment,
						$ExRoleStrings,
						$Pre2007 = $False)
					if ($Pre2007)
					{
						$BGColHeader = "#880099"
						$BGColSubHeader = "#8800CC"
						$Prefix = ""
						$IntNamesText = ""
						$ExtNamesText = ""
						$CASArrayText = ""
					}
					else
					{
						$BGColHeader = "#000099"
						$BGColSubHeader = "#0000FF"
						$Prefix = "Site:"
						$IntNamesText = ""
						$ExtNamesText = ""
						$CASArrayText = ""
						$IntNames = @()
						$ExtNames = @()
						$CASArrayName = ""
						foreach ($Server in $Servers.Value)
						{
							$IntNames += $Server.IntNames
							$ExtNames += $Server.ExtNames
							$CASArrayName = $Server.CASArrayName
							
						}
						$IntNames = $IntNames | Sort -Unique
						$ExtNames = $ExtNames | Sort -Unique
						$IntNames = [system.String]::Join(",", $IntNames)
						$ExtNames = [system.String]::Join(",", $ExtNames)
						if ($IntNames)
						{
							$IntNamesText = "Internal Names: $($IntNames)"
							$ExtNamesText = "External Names: $($ExtNames)<br >"
						}
						if ($CASArrayName)
						{
							$CASArrayText = "CAS Array: $($CASArrayName)"
						}
					}
					$Output = "<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Segoe UI,Arial,sans-serif"">
       <col width=""20%""><col width=""20%"">
       <colgroup width=""25%"">";
					
					$ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort Name | %{ $Output += "<col width=""3%"">" }
					$Output += "</colgroup><col width=""20%""><col  width=""20%"">
       <tr bgcolor=""$($BGColHeader)""><th><font color=""#ffffff"">$($Prefix) $($Servers.Key)</font></th>
       <th colspan=""$(($ExchangeEnvironment.TotalServersByRole.Count) + 2)"" align=""left""><font color=""#ffffff"">$($ExtNamesText)$($IntNamesText)</font></th>
       <th align=""center""><font color=""#ffffff"">$($CASArrayText)</font></th></tr>"
					$TotalMailboxes = 0
					$Servers.Value | %{ $TotalMailboxes += $_.Mailboxes }
					$Output += "<tr bgcolor=""$($BGColSubHeader)""><th><font color=""#ffffff"">Mailboxes: $($TotalMailboxes)</font></th><th>"
					$Output += "<font color=""#ffffff"">Exchange Version</font></th>"
					$ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort Name | %{ $Output += "<th><font color=""#ffffff"">$($ExRoleStrings[$_.Key].Short)</font></th>" }
					$Output += "<th><font color=""#ffffff"">OS Version</font></th><th><font color=""#ffffff"">OS Service Pack</font></th></tr>"
					$AlternateRow = 0
					
					foreach ($Server in $Servers.Value)
					{
						$Output += "<tr "
						if ($AlternateRow)
						{
							$Output += " style=""background-color:#dddddd"""
							$AlternateRow = 0
						}
						else
						{
							$AlternateRow = 1
						}
						$Output += "><td>$($Server.Name)"
						if ($Server.RealName -ne $Server.Name)
						{
							$Output += " ($($Server.RealName))"
						}
						$Output += "</td><td>$($ExVersionStrings["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].Long)"
						if ($Server.RollupLevel -gt 0)
						{
							$Output += " UR$($Server.RollupLevel)"
							if ($Server.RollupVersion)
							{
								$Output += " $($Server.RollupVersion)"
							}
						}
						$Output += "</td>"
						$ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort Name | %{
							$Output += "<td"
							if ($Server.Roles -contains $_.Key)
							{
								$Output += " align=""center"" style=""background-color:#00FF00"""
							}
							$Output += ">"
							if (($_.Key -eq "ClusteredMailbox" -or $_.Key -eq "Mailbox" -or $_.Key -eq "BE") -and $Server.Roles -contains $_.Key)
							{
								$Output += $Server.Mailboxes
							}
						}
						
						$Output += "<td>$($Server.OSVersion)</td><td>$($Server.OSServicePack)</td></tr>";
					}
					$Output += "<tr></tr>
       </table><br />"
					$Output
				}
				
				# Sub Function to return HTML Table for Databases
				function _GetDBTable
				{
					param ($Databases)
					# Only Show Archive Mailbox Columns, Backup Columns and Circ Logging if at least one DB has an Archive mailbox, backed up or Cir Log enabled.
					$ShowArchiveDBs = $False
					$ShowLastFullBackup = $False
					$ShowCircularLogging = $False
					$ShowStorageGroups = $False
					$ShowCopies = $False
					$ShowFreeDatabaseSpace = $False
					$ShowFreeLogDiskSpace = $False
					foreach ($Database in $Databases)
					{
						if ($Database.ArchiveMailboxCount -gt 0)
						{
							$ShowArchiveDBs = $True
						}
						if ($Database.LastFullBackup -ne "Not Available")
						{
							$ShowLastFullBackup = $True
						}
						if ($Database.CircularLoggingEnabled -eq "Yes")
						{
							$ShowCircularLogging = $True
						}
						if ($Database.StorageGroup)
						{
							$ShowStorageGroups = $True
						}
						if ($Database.CopyCount -gt 0)
						{
							$ShowCopies = $True
						}
						if ($Database.FreeDatabaseDiskSpace -ne $null)
						{
							$ShowFreeDatabaseSpace = $true
						}
						if ($Database.FreeLogDiskSpace -ne $null)
						{
							$ShowFreeLogDiskSpace = $true
						}
					}
					
					
					$Output = "<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Segoe UI,Arial,sans-serif"">
       
       <tr align=""center"" bgcolor=""#FFD700"">
       <th>Server</th>"
					if ($ShowStorageGroups)
					{
						$Output += "<th>Storage Group</th>"
					}
					$Output += "<th>Database Name</th>
       <th>Mailboxes</th>
       <th>Av. Mailbox Size</th>"
					if ($ShowArchiveDBs)
					{
						$Output += "<th>Archive MBs</th><th>Av. Archive Size</th>"
					}
					$Output += "<th>DB Size</th><th>DB Whitespace</th>"
					if ($ShowFreeDatabaseSpace)
					{
						$Output += "<th>Database Disk Free</th>"
					}
					if ($ShowFreeLogDiskSpace)
					{
						$Output += "<th>Log Disk Free</th>"
					}
					if ($ShowLastFullBackup)
					{
						$Output += "<th>Last Full Backup</th>"
					}
					if ($ShowCircularLogging)
					{
						$Output += "<th>Circular Logging</th>"
					}
					if ($ShowCopies)
					{
						$Output += "<th>Copies (n)</th>"
					}
					
					$Output += "</tr>"
					$AlternateRow = 0;
					foreach ($Database in $Databases)
					{
						$Output += "<tr"
						if ($AlternateRow)
						{
							$Output += " style=""background-color:#dddddd"""
							$AlternateRow = 0
						}
						else
						{
							$AlternateRow = 1
						}
						
						$Output += "><td>$($Database.ActiveOwner)</td>"
						if ($ShowStorageGroups)
						{
							$Output += "<td>$($Database.StorageGroup)</td>"
						}
						$Output += "<td>$($Database.Name)</td>
             <td align=""center"">$($Database.MailboxCount)</td>
             <td align=""center"">$("{0:N2}" -f ($Database.MailboxAverageSize/1MB)) MB</td>"
						if ($ShowArchiveDBs)
						{
							$Output += "<td align=""center"">$($Database.ArchiveMailboxCount)</td> 
                    <td align=""center"">$("{0:N2}" -f ($Database.ArchiveAverageSize/1MB)) MB</td>";
						}
						$Output += "<td align=""center"">$("{0:N2}" -f ($Database.Size/1GB)) GB </td>
             <td align=""center"">$("{0:N2}" -f ($Database.Whitespace/1GB)) GB</td>";
						if ($ShowFreeDatabaseSpace)
						{
							$Output += "<td align=""center"">$("{0:N1}" -f $Database.FreeDatabaseDiskSpace)%</td>"
						}
						if ($ShowFreeLogDiskSpace)
						{
							$Output += "<td align=""center"">$("{0:N1}" -f $Database.FreeLogDiskSpace)%</td>"
						}
						if ($ShowLastFullBackup)
						{
							$Output += "<td align=""center"">$($Database.LastFullBackup)</td>";
						}
						if ($ShowCircularLogging)
						{
							$Output += "<td align=""center"">$($Database.CircularLoggingEnabled)</td>";
						}
						if ($ShowCopies)
						{
							$Output += "<td>$($Database.Copies | %{ $_ }) ($($Database.CopyCount))</td>"
						}
						$Output += "</tr>";
					}
					$Output += "</table><br />"
					
					$Output
				}
				
				
				# Sub Function to neatly update progress
				function _UpProg1
				{
					param ($PercentComplete,
						$Status,
						$Stage)
					$TotalStages = 5
					Write-Progress -id 1 -activity "Get-ExchangeEnvironmentReport" -status $Status -percentComplete (($PercentComplete/$TotalStages) + (1/$TotalStages * $Stage * 100))
				}
				
				# 1. Initial Startup
				
				# 1.0 Check Powershell Version
				if ((Get-Host).Version.Major -eq 1)
				{
					throw "Powershell Version 1 not supported";
				}
				
				# 1.1 Check Exchange Management Shell, attempt to load
				if (!(Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue))
				{
					if (Test-Path "C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1")
					{
						. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'
						Connect-ExchangeServer -auto
					}
					elseif (Test-Path "C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1")
					{
						Add-PSSnapIn Microsoft.Exchange.Management.PowerShell.Admin
						.'C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1'
					}
					else
					{
						throw "Exchange Management Shell cannot be loaded"
					}
				}
				
				# 1.2 Check if -SendMail parameter set and if so check -MailFrom, -MailTo and -MailServer are set
				if ($SendMail)
				{
					if (!$MailFrom -or !$MailTo -or !$MailServer)
					{
						throw "If -SendMail specified, you must also specify -MailFrom, -MailTo and -MailServer"
					}
				}
				
				# 1.3 Check Exchange Management Shell Version
				if ((Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue))
				{
					$E2010 = $false;
					if (Get-ExchangeServer | Where { $_.AdminDisplayVersion.Major -gt 14 })
					{
						Write-Warning "Exchange 2010 or higher detected. You'll get better results if you run this script from the latest management shell"
					}
				}
				else
				{
					
					$E2010 = $true
					$localserver = get-exchangeserver $Env:computername
					$localversion = $localserver.admindisplayversion.major
					if ($localversion -eq 15) { $E2013 = $true }
					
				}
				
				# 1.4 Check view entire forest if set (by default, true)
				if ($E2010)
				{
					Set-ADServerSettings -ViewEntireForest:$ViewEntireForest
				}
				else
				{
					$global:AdminSessionADSettings.ViewEntireForest = $ViewEntireForest
				}
				
				# 1.5 Initial Variables
				
				# 1.5.1 Hashtable to update with environment data
				$ExchangeEnvironment = @{
					Sites = @{ }
					Pre2007 = @{ }
					Servers = @{ }
					DAGs  = @()
					NonDAGDatabases = @()
				}
				# 1.5.7 Exchange Major Version String Mapping
				$ExMajorVersionStrings = @{
					"6.0" = @{ Long = "Exchange 2000"; Short = "E2000" }
					"6.5" = @{ Long = "Exchange 2003"; Short = "E2003" }
					"8"   = @{ Long = "Exchange 2007"; Short = "E2007" }
					"14"  = @{ Long = "Exchange 2010"; Short = "E2010" }
					"15"  = @{ Long = "Exchange 2013"; Short = "E2013" }
					"15.1" = @{ Long = "Exchange 2016"; Short = "E2016" }
				}
				# 1.5.8 Exchange Service Pack String Mapping
				$ExSPLevelStrings = @{
					"0"   = "RTM"
					"1"   = "SP1"
					"2"   = "SP2"
					"3"   = "SP3"
					"4"   = "SP4"
					"SP1" = "SP1"
					"SP2" = "SP2"
				}
				# Add many CUs               
				for ($i = 1; $i -le 20; $i++)
				{
					$ExSPLevelStrings.Add("CU$($i)", "CU$($i)");
				}
				# 1.5.9 Populate Full Mapping using above info
				$ExVersionStrings = @{ }
				foreach ($Major in $ExMajorVersionStrings.GetEnumerator())
				{
					foreach ($Minor in $ExSPLevelStrings.GetEnumerator())
					{
						$ExVersionStrings.Add("$($Major.Key).$($Minor.Key)", @{ Long = "$($Major.Value.Long) $($Minor.Value)"; Short = "$($Major.Value.Short)$($Minor.Value)" })
					}
				}
				# 1.5.10 Exchange Role String Mapping
				$ExRoleStrings = @{
					"ClusteredMailbox" = @{ Short = "ClusMBX"; Long = "CCR/SCC Clustered Mailbox" }
					"Mailbox"		   = @{ Short = "MBX"; Long = "Mailbox" }
					"ClientAccess"	   = @{ Short = "CAS"; Long = "Client Access" }
					"HubTransport"	   = @{ Short = "HUB"; Long = "Hub Transport" }
					"UnifiedMessaging" = @{ Short = "UM"; Long = "Unified Messaging" }
					"Edge"			   = @{ Short = "EDGE"; Long = "Edge Transport" }
					"FE"			   = @{ Short = "FE"; Long = "Front End" }
					"BE"			   = @{ Short = "BE"; Long = "Back End" }
					"Hybrid"		   = @{ Short = "HYB"; Long = "Hybrid" }
					"Unknown"		   = @{ Short = "Unknown"; Long = "Unknown" }
				}
				
				# 2 Get Relevant Exchange Information Up-Front
				
				# 2.1 Get Server, Exchange and Mailbox Information
				_UpProg1 1 "Getting Exchange Server List" 1
				$ExchangeServers = [array](Get-ExchangeServer $ServerFilter)
				if (!$ExchangeServers)
				{
					throw "No Exchange Servers matched by -ServerFilter ""$($ServerFilter)"""
				}
				$HybridServers = @()
				if (Get-Command Get-HybridConfiguration -ErrorAction SilentlyContinue)
				{
					$HybridConfig = Get-HybridConfiguration
					$HybridConfig.ReceivingTransportServers | %{ $HybridServers += $_.Name }
					$HybridConfig.SendingTransportServers | %{ $HybridServers += $_.Name }
					$HybridServers = $HybridServers | Sort-Object -Unique
				}
				
				_UpProg1 10 "Getting Mailboxes" 1
				$Mailboxes = [array](Get-Mailbox -ResultSize Unlimited) | Where { $_.Server -like $ServerFilter }
				if ($E2010)
				{
					_UpProg1 60 "Getting Archive Mailboxes" 1
					$ArchiveMailboxes = [array](Get-Mailbox -Archive -ResultSize Unlimited) | Where { $_.Server -like $ServerFilter }
					_UpProg1 70 "Getting Remote Mailboxes" 1
					$RemoteMailboxes = [array](Get-RemoteMailbox -ResultSize Unlimited)
					$ExchangeEnvironment.Add("RemoteMailboxes", $RemoteMailboxes.Count)
					_UpProg1 90 "Getting Databases" 1
					if ($E2013)
					{
						$Databases = [array](Get-MailboxDatabase -IncludePreExchange2013 -Status) | Where { $_.Server -like $ServerFilter }
					}
					elseif ($E2010)
					{
						$Databases = [array](Get-MailboxDatabase -IncludePreExchange2010 -Status) | Where { $_.Server -like $ServerFilter }
					}
					$DAGs = [array](Get-DatabaseAvailabilityGroup) | Where { $_.Servers -like $ServerFilter }
				}
				else
				{
					$ArchiveMailboxes = $null
					$ArchiveMailboxStats = $null
					$DAGs = $null
					_UpProg1 90 "Getting Databases" 1
					$Databases = [array](Get-MailboxDatabase -IncludePreExchange2007 -Status) | Where { $_.Server -like $ServerFilter }
					$ExchangeEnvironment.Add("RemoteMailboxes", 0)
				}
				
				# 2.3 Populate Information we know
				$ExchangeEnvironment.Add("TotalMailboxes", $Mailboxes.Count + $ExchangeEnvironment.RemoteMailboxes);
				
				# 3 Process High-Level Exchange Information
				
				# 3.1 Collect Exchange Server Information
				for ($i = 0; $i -lt $ExchangeServers.Count; $i++)
				{
					_UpProg1 ($i/$ExchangeServers.Count * 100) "Getting Exchange Server Information" 2
					# Get Exchange Info
					$ExSvr = _GetExSvr -E2010 $E2010 -ExchangeServer $ExchangeServers[$i] -Mailboxes $Mailboxes -Databases $Databases -Hybrids $HybridServers
					# Add to site or pre-Exchange 2007 list
					if ($ExSvr.Site)
					{
						# Exchange 2007 or higher
						if (!$ExchangeEnvironment.Sites[$ExSvr.Site])
						{
							$ExchangeEnvironment.Sites.Add($ExSvr.Site, @($ExSvr))
						}
						else
						{
							$ExchangeEnvironment.Sites[$ExSvr.Site] += $ExSvr
						}
					}
					else
					{
						# Exchange 2003 or lower
						if (!$ExchangeEnvironment.Pre2007["Pre 2007 Servers"])
						{
							$ExchangeEnvironment.Pre2007.Add("Pre 2007 Servers", @($ExSvr))
						}
						else
						{
							$ExchangeEnvironment.Pre2007["Pre 2007 Servers"] += $ExSvr
						}
					}
					# Add to Servers List
					$ExchangeEnvironment.Servers.Add($ExSvr.Name, $ExSvr)
				}
				
				# 3.2 Calculate Environment Totals for Version/Role using collected data
				_UpProg1 1 "Getting Totals" 3
				$ExchangeEnvironment.Add("TotalMailboxesByVersion", (_TotalsByVersion -ExchangeEnvironment $ExchangeEnvironment))
				$ExchangeEnvironment.Add("TotalServersByRole", (_TotalsByRole -ExchangeEnvironment $ExchangeEnvironment))
				
				# 3.4 Populate Environment DAGs
				_UpProg1 5 "Getting DAG Info" 3
				if ($DAGs)
				{
					foreach ($DAG in $DAGs)
					{
						$ExchangeEnvironment.DAGs += (_GetDAG -DAG $DAG)
					}
				}
				
				# 3.5 Get Database information
				_UpProg1 60 "Getting Database Info" 3
				for ($i = 0; $i -lt $Databases.Count; $i++)
				{
					$Database = _GetDB -Database $Databases[$i] -ExchangeEnvironment $ExchangeEnvironment -Mailboxes $Mailboxes -ArchiveMailboxes $ArchiveMailboxes -E2010 $E2010
					$DAGDB = $false
					for ($j = 0; $j -lt $ExchangeEnvironment.DAGs.Count; $j++)
					{
						if ($ExchangeEnvironment.DAGs[$j].Members -contains $Database.ActiveOwner)
						{
							$DAGDB = $true
							$ExchangeEnvironment.DAGs[$j].Databases += $Database
						}
					}
					if (!$DAGDB)
					{
						$ExchangeEnvironment.NonDAGDatabases += $Database
					}
					
					
				}
				
				# 4 Write Information
				_UpProg1 5 "Writing HTML Report Header" 4
				# Header
				$Output = "<html>
<body>
<font size=""1"" face=""Segoe UI,Arial,sans-serif"">
<h2 align=""center"">Exchange Environment Report</h3>
<h4 align=""center"">Generated $((Get-Date).ToString())</h5>
</font>
<table border=""0"" cellpadding=""3"" style=""font-size:8pt;font-family:Segoe UI,Arial,sans-serif"">
<tr bgcolor=""#009900"">
<th colspan=""$($ExchangeEnvironment.TotalMailboxesByVersion.Count)""><font color=""#ffffff"">Total Servers:</font></th>"
				if ($ExchangeEnvironment.RemoteMailboxes)
				{
					$Output += "<th colspan=""$($ExchangeEnvironment.TotalMailboxesByVersion.Count + 2)""><font color=""#ffffff"">Total Mailboxes:</font></th>"
				}
				else
				{
					$Output += "<th colspan=""$($ExchangeEnvironment.TotalMailboxesByVersion.Count + 1)""><font color=""#ffffff"">Total Mailboxes:</font></th>"
				}
				$Output += "<th colspan=""$($ExchangeEnvironment.TotalServersByRole.Count)""><font color=""#ffffff"">Total Roles:</font></th></tr>
<tr bgcolor=""#00CC00"">"
				# Show Column Headings based on the Exchange versions we have
				$ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort Name | %{ $Output += "<th>$($ExVersionStrings[$_.Key].Short)</th>" }
				$ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort Name | %{ $Output += "<th>$($ExVersionStrings[$_.Key].Short)</th>" }
				if ($ExchangeEnvironment.RemoteMailboxes)
				{
					$Output += "<th>Office 365</th>"
				}
				$Output += "<th>Org</th>"
				$ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort Name | %{ $Output += "<th>$($ExRoleStrings[$_.Key].Short)</th>" }
				$Output += "<tr>"
				$Output += "<tr align=""center"" bgcolor=""#dddddd"">"
				$ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort Name | %{ $Output += "<td>$($_.Value.ServerCount)</td>" }
				$ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort Name | %{ $Output += "<td>$($_.Value.MailboxCount)</td>" }
				if ($RemoteMailboxes)
				{
					$Output += "<th>$($ExchangeEnvironment.RemoteMailboxes)</th>"
				}
				$Output += "<td>$($ExchangeEnvironment.TotalMailboxes)</td>"
				$ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort Name | %{ $Output += "<td>$($_.Value)</td>" }
				$Output += "</tr><tr><tr></table><br>"
				
				# Sites and Servers
				_UpProg1 20 "Writing HTML Site Information" 4
				foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator())
				{
					$Output += _GetOverview -Servers $Site -ExchangeEnvironment $ExchangeEnvironment -ExRoleStrings $ExRoleStrings
				}
				_UpProg1 40 "Writing HTML Pre-2007 Information" 4
				foreach ($FakeSite in $ExchangeEnvironment.Pre2007.GetEnumerator())
				{
					$Output += _GetOverview -Servers $FakeSite -ExchangeEnvironment $ExchangeEnvironment -ExRoleStrings $ExRoleStrings -Pre2007:$true
				}
				
				_UpProg1 60 "Writing HTML DAG Information" 4
				foreach ($DAG in $ExchangeEnvironment.DAGs)
				{
					if ($DAG.MemberCount -gt 0)
					{
						# Database Availability Group Header
						$Output += "<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Segoe UI,Arial,sans-serif"">
             <col width=""20%""><col width=""10%""><col width=""70%"">
             <tr align=""center"" bgcolor=""#FF8000 ""><th>Database Availability Group Name</th><th>Member Count</th>
             <th>Database Availability Group Members</th></tr>
             <tr><td>$($DAG.Name)</td><td align=""center"">
             $($DAG.MemberCount)</td><td>"
						$DAG.Members | % { $Output += "$($_) " }
						$Output += "</td></tr></table>"
						
						# Get Table HTML
						$Output += _GetDBTable -Databases $DAG.Databases
					}
					
				}
				
				if ($ExchangeEnvironment.NonDAGDatabases.Count)
				{
					_UpProg1 80 "Writing HTML Non-DAG Database Information" 4
					$Output += "<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Segoe UI,Arial,sans-serif"">
         <tr bgcolor=""#FF8000""><th>Mailbox Databases (Non-DAG)</th></table>"
					$Output += _GetDBTable -Databases $ExchangeEnvironment.NonDAGDatabases
				}
				
				
				# End
				_UpProg1 90 "Finishing off.." 4
				$Output += "</body></html>";
				$Output | Out-File $HTMLReport
				
				
				if ($SendMail)
				{
					_UpProg1 95 "Sending mail message.." 4
					Send-MailMessage -Attachments $HTMLReport -To $MailTo -From $MailFrom -Subject "Exchange Environment Report" -BodyAsHtml $Output -SmtpServer $MailServer
				}
				
				$Reboot = $true
			}#endregion
			#region Option 62) Generate Mailbox Size and Information Reports
			62 {
				"This function is not yet implemented"
				#      Generate Mailbox Size and Information Reports
				<#ModuleStatus -name ServerManager
				Install-WinUniComm4
				$Reboot = $false#>
			}
			#endregion
			#region Option 63) Generate Reports for Exchange ActiveSync Device Statistics
			63 {
				generateEASDeviceStats
			}
			#endregion
			#region Option 64) Exchange Analyzer
			64 {
				"This function is not yet implemented"
				#      Exchange Analyzer
				<#empty
				$Reboot = $false#>
			}
			#endregion
			#region Option 65) Generate Report Total Emails Sent and Received Per Day and Size
			65 {
				#      Generate Report Total Emails Sent and Received Per Day and Size
				# Script:    TotalEmailsSentReceivedPerDay.ps1
				# Purpose:   Get the number of e-mails sent and received per day
				# Author:    Nuno Mota
				# Date:             October 2010
				#region user input
				"Get the number of e-mails sent and received per day"
				"Enter start date"
				[INT]$MM = Read-Host "Month"
				[INT]$DD = Read-Host "Day"
				[INT]$YY = Read-Host "Year"
				
				[INT]$noOfdays = Read-Host "Enter number of days. (Start date included)"
				#endregion
				
				# Initialize date variables used for counting and for output 
				$From = Get-Date -Month $MM -Day $DD -Year $YY
				$To = $From.AddDays(1)
				$End = $from.AddDays($noOfdays)
				
				[Int64]$intSent = $intRec = 0
				[Int64]$intSentSize = $intRecSize = 0
				[String]$strEmails = $null
				
				Write-Host "DayOfWeek,Date,Sent,Sent Size (MB),Received,Received Size (MB)" -ForegroundColor Yellow
				
				Do
				{
					# Start building the variable that will hold the information for the day 
					$strEmails = "$($From.DayOfWeek),$($From.ToShortDateString()),"
					
					$intSent = $intRec = 0
					(Get-TransportService) | Get-MessageTrackingLog -ResultSize Unlimited -Start $From -End $To | ForEach {
						# Sent E-mails 
						If ($_.EventId -eq "RECEIVE" -and $_.Source -eq "STOREDRIVER")
						{
							$intSent++
							$intSentSize += $_.TotalBytes
						}
						
						# Received E-mails 
						If ($_.EventId -eq "DELIVER")
						{
							$intRec += $_.RecipientCount
							$intRecSize += $_.TotalBytes
						}
					}
					
					$intSentSize = [Math]::Round($intSentSize/1MB, 0)
					$intRecSize = [Math]::Round($intRecSize/1MB, 0)
					
					# Add the numbers to the $strEmails variable and print the result for the day 
					$strEmails += "$intSent,$intSentSize,$intRec,$intRecSize"
					$strEmails
					
					# Increment the From and To by one day 
					$From = $From.AddDays(1)
					$To = $From.AddDays(1)
				}
				Until ($From -eq $end)
				$Reboot = $false
				
			}
			#endregion
			#region Option 66) Generate HTML Report for Mailbox Permissions
			66 {
				#      Generate HTML Report for Mailbox Permissions
                           <#
    .SYNOPSIS
    Creates a HTML Report showing Sendas, Full Access and Send on Behalf Permission Information for Each Mailbox for your Exchange Organization, selected database or for a single user.
   
       Serkan Varoglu
       
       http:\\Mshowto.org
       http:\\Get-Mailbox.org
       
       THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
       RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
       
       Version 1.1, 5 March 2012
       
    .DESCRIPTION
       
    Creates a HTML Report showing Sendas, Full Access and Send on Behalf Permission Information for Each Mailbox for your Exchange Organization, selected database or for a single user.
       By Default Inherited Send As permission and NT Authority\Self account will not be shown in the report unless you run the script with the parameters listed below.
       Also by default all mailboxes will be reported if you want to report a single database, you can use -database parameter to specify your database name or you can get the report for a single user.
       
       .PARAMETER HTMLReport
    Filename to write HTML Report to
       
       .PARAMETER Database
    By default this script will report all mailboxes. If you want to report mailboxes in a single database, you can use this parameter to input your database name.
       
       .PARAMETER Mailbox
    By default this script will report all mailboxes. If you want to report a single mailbox, you can use this parameter to input the mailbox you want to report.
       
       .SWITCH ShowInherited
       If ShowInherited is added as switch the report will show Inherited Sendas permissions for mailboxes as well.
       
       .SWITCH ShowSelf
       If ShowSelf is added as switch the report will show "NT Authority\Self" sendas permission for mailboxes as well.
       
       .EXAMPLE
    Generate the HTML report 
    .\Report-Permissions.ps1 -HTMLReport "C:\Users\SVaroglu\Desktop\MailboxPermissionReport.HTML"
       
#>
				#region user input
				$HTMLReport = Read-Host "Specifiy alternate path and name for report file. Default is ""mailboxpermissionsreport.html"""
				if ($HTMLReport = "")
				{
					$ReportFile = "mailboxpermissionsreport.html"
				}
				$ShowInheritedYN = Read-Host "List inherited SendAs and Full Access permissions?[Y/N] Default is [Y]"
				switch ($ShowInheritedYN)
				{
					Y{ $ShowInherited = $true }
					N{ $ShowInherited = $false }
					default { $ShowInherited = $true }
				}
				$ShowSelfYN = Read-Host "List NT Authority\Self Permission ?[Y/N] Default is [Y]"
				switch ($ShowInheritedYN)
				{
					Y{ $ShowSelf = $true }
					N{ $ShowSelf = $false }
					default { $ShowSelf = $true }
				}
				$MailboxYN = Read-Host "Specify a mailbox to report?[Y/N] Default is [N]"
				switch ($MailboxYN)
				{
					Y{ $Mailbox = Read-Host "Enter mailbox name" }
					N{ $Mailbox = $null }
					default { $Mailbox = $null }
				}
				#endregion
				$Watch = [System.Diagnostics.Stopwatch]::StartNew()
				$WarningPreference = "SilentlyContinue"
				$ErrorActionPreference = "SilentlyContinue"
				$ShowInherited = $ShowInherited.IsPresent
				$ShowSelf = $ShowSelf.IsPresent
				$u = 1
				$s = 0
				$f = 0
				$b = 0
				$n = 0
				$nj = -1
				$gj = -1
				if (!$database) { $dbnull = 0 }
				if (!$mailbox) { $mbnull = 0 }
				if (($ShowSelf -like "false") -and ($ShowInherited -like "false"))
				{ $gentitle = "Mailboxes With Custom Permissions" }
				else
				{ $gentitle = "Mailboxes" }
				$gen = "<table border=""1"" bordercolor=""#4384D3"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#4384D3"" align=""center""><th colspan=""5""><font color=""#FFFFFF"">$($gentitle)</font></th></tr><tr>"
				$inh = "<table border=""1"" bordercolor=""#4384D3"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#4384D3"" align=""center""><th colspan=""5""><font color=""#FFFFFF"">Mailboxes With Only Inherited Permissions</font></th></tr><tr>"
				function _Progress
				{
					param ($PercentComplete,
						$Status)
					Write-Progress -id 1 -activity "Report for Mailboxes" -status $Status -percentComplete ($PercentComplete)
				}
				_Progress (($u * 100)/100) "Collecting Mailbox Information"
				if (!$database -and !$mailbox)
				{
					$mailboxes = get-mailbox -resultsize unlimited | Sort-Object name
				}
				elseif ($database -and !$mailbox)
				{
					$mailboxes = get-mailbox -database $database -resultsize unlimited | Sort-Object name
				}
				elseif (!$database -and $mailbox)
				{
					$mailboxes = get-mailbox $mailbox
				}
				else
				{
					Write-Host -ForegroundColor Cyan "Please choose database or single mailbox. Both Parameters can not be used at the same time. Ended without compiling a report."
					exit
				}
				$mcount = ($mailboxes | measure-object).count
				if ($mcount -eq 0)
				{
					Write-Host -ForegroundColor Cyan "No Mailbox Found. Ended without compiling a report. Please Check Your Input."
					exit
				}
				foreach ($mailbox in $mailboxes)
				{
					_Progress (($u * 95)/$mcount) "Processing $mailbox, $($u) of $($mcount) Mailboxes."
					$SenderBody = ""
					$FullBody = ""
					$BehalfBody = ""
					$sendbehalfs = Get-Mailbox $mailbox | select-object -expand grantsendonbehalfto | select-object -expand rdn | Sort-Object Unescapedname
					if (($ShowSelf -like "true") -and ($ShowInherited -like "true"))
					{
						$senders = Get-ADPermission $mailbox.identity | ?{ ($_.extendedrights -like "*send-as*") } | Sort-Object name
						$fullsenders = Get-Mailbox $mailbox | Get-MailboxPermission | ?{ ($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") }
					}
					elseif (($ShowSelf -like "false") -and ($ShowInherited -like "true"))
					{
						$senders = Get-ADPermission $mailbox.identity | ?{ ($_.extendedrights -like "*send-as*") -and ($_.User -notlike "NT Authority\self") } | Sort-Object name
						$fullsenders = Get-Mailbox $mailbox | Get-MailboxPermission | ?{ ($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\self*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") }
					}
					elseif (($ShowSelf -like "true") -and ($ShowInherited -like "false"))
					{
						$senders = Get-ADPermission $mailbox.identity | ?{ ($_.extendedrights -like "*send-as*") -and ($_.isinherited -like "false") } | Sort-Object name
						$fullsenders = Get-Mailbox $mailbox | Get-MailboxPermission | ?{ ($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") -and ($_.IsInherited -like "false") }
					}
					else
					{
						$senders = Get-ADPermission $mailbox.identity | ?{ ($_.extendedrights -like "*send-as*") -and ($_.isinherited -like "false") -and ($_.User -notlike "NT Authority\self") } | Sort-Object name
						$fullsenders = Get-Mailbox $mailbox | Get-MailboxPermission | ?{ ($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\self*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") -and ($_.IsInherited -like "false") }
					}
					if (!$senders -and !$fullsenders -and !$sendbehalfs)
					{
						$n++
						if ($nj -eq 4)
						{
							$inh += "</tr><tr><td>$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</td>"
							$nj = 0
						}
						else
						{
							$inh += "<td>$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</td>"
							$nj++
						}
					}
					else
					{
						if ($gj -eq 4)
						{
							$gen += "</tr><tr><td><a href=""#$($mailbox.Name)"">$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</a></td>"
							$gj = 0
						}
						else
						{
							$gen += "<td><a href=""#$($mailbox.Name)"">$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</a></td>"
							$gj++
						}
						$MailboxTable = "<table border=""1"" bordercolor=""#1F497B"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#1F497B"" align=""center""><th colspan=""3"" ><font color=""#FFFFFF""><a name=""$($mailbox.Name)"">$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</font></a></th></tr><tr>"
						if (!$senders)
						{
							$SenderBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#696969"" align=""center"" valign=""top""><font color=""#FFFF99"">No User has Send As Permission On This Mailbox</font></td></table></td>"
						}
						else
						{
							$SenderBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><tr><td colspan=""3"" align=""center"" valign=""top"" bgcolor=""#696969""><font color=""#FFFFE0"">Send-As Permissions</font></td></tr><tr bgcolor=""#878787"" align=""center"">
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Send as Permission Owner</font></td>
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Deny</font></td>
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Inherited</font></td>
                                        </tr>"
							foreach ($sender in $senders)
							{
								if (0, 2, 4, 6, 8 -contains "$sj"[-1] - 48)
								{
									$bgcolor = "'#E8E8E8'"
								}
								else
								{
									$bgcolor = "'#C8C8C8'"
								}
								$SenderBody += "<tr align=""center"" bgcolor=$($bgcolor)>"
								$SenderBody += "<td><font color=""#003333"">$($sender.user)</font></td>"
								if ($sender.deny -like "true") { $font = "red" }
								else { $font = "'#000000'" }
								$SenderBody += "<td><font color=$font>$($sender.deny)</font></td>"
								if ($sender.isinherited -like "false") { $font = "red" }
								else { $font = "'#000000'" }
								$SenderBody += "<td><font color=$font>$($sender.isinherited)</font></td>"
								$SenderBody += "</tr>"
								$sj++
							}
							$SenderBody += "</table></td>"
							$s++
						}
						
						if (!$fullsenders)
						{
							$FullBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#696969"" align=""center"" valign=""top""><font color=""#FFFF99"">No User has Full Access On This Mailbox</font></td></table></td>"
						}
						else
						{
							$FullBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><tr><td colspan=""3"" align=""center"" valign=""top"" bgcolor=""#696969""><font color=""#FFFFE0"">Full Access Permissions</font></td></tr><tr bgcolor=""#878787"" align=""center"">
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Full Access Permission Owner</font></td>
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Deny</font></td>
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Inherited</font></td>
                                        </tr>"
							foreach ($fullsender in $fullsenders)
							{
								if (0, 2, 4, 6, 8 -contains "$fj"[-1] - 48)
								{
									$bgcolor = "'#E8E8E8'"
								}
								else
								{
									$bgcolor = "'#C8C8C8'"
								}
								$FullBody += "<tr align=""center"" bgcolor=$($bgcolor)>"
								$FullBody += "<td><font color=""#003333"">$($fullsender.user)</font></td>"
								if ($fullsender.deny -like "true") { $font = "red" }
								else { $font = "'#000000'" }
								$FullBody += "<td><font color=$font>$($fullsender.deny)</font></td>"
								if ($fullsender.isinherited -like "false") { $font = "red" }
								else { $font = "'#000000'" }
								$FullBody += "<td><font color=$font>$($fullsender.isinherited)</font></td>"
								$FullBody += "</tr>"
								$fj++
							}
							$FullBody += "</table></td>"
							$f++
						}
						
						if (!$sendbehalfs)
						{
							$BehalfBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#696969"" align=""center"" valign=""top""><font color=""#FFFF99"">No User has Send on Behalf On This Mailbox</font></td></table></td>"
						}
						else
						{
							$BehalfBody += "<td align=""center"" valign=""top"" width=""33%"">
                                        <table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
                                        <tr><td align=""center valign=""top"" bgcolor=""#696969""><font color=""#FFFFE0"">Send on Behalf</font></td></tr>
                                        <tr><td bgcolor=""#878787"" nowrap=""nowrap""><font color=""#FFFFFF"">Send On Behalf Permission Owner</font></td></tr>"
							foreach ($sendbehalf in $sendbehalfs)
							{
								if (0, 2, 4, 6, 8 -contains "$bj"[-1] - 48)
								{
									$bgcolor = "'#E8E8E8'"
								}
								else
								{
									$bgcolor = "'#C8C8C8'"
								}
								$BehalfBody += "<tr align=""center"" bgcolor=$($bgcolor)>"
								$BehalfBody += "<td><font color=""#003333"">$($sendbehalf.unescapedname)</font></td>"
								$BehalfBody += "</tr>"
								$bj++
							}
							$BehalfBody += "</table></td>"
							$b++
						}
						$Table += $MailboxTable + $SenderBody + $FullBody + $BehalfBody + "</tr></table><br><a href=""#top"">&#9650;</a><hr /><br>"
					}
					$u++
				}
				_Progress (98) "Completing"
				if (($ShowSelf -like "false") -and ($ShowInherited -like "false"))
				{
					if (($dbnull -eq 0) -and ($mbnull -eq 0))
					{
						$Summary = "<table style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#FFE87C"" >In your Exchange Organization there are $($mcount) mailboxes present."
						$Summary += "Send as Permission explicity configured on $($s) of these mailboxes. Full Access Permission explicity configured on $($f) of these mailboxes. Send on Behalf explicity configured on $($b) of these mailboxes and $($n) mailbox has inherited permissions only.<td></table><br>"
					}
					elseif (($dbnull -ne 0) -and ($mbnull -eq 0))
					{
						$Summary = "<table style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#FFE87C"" >In $($database) mailbox database, there are $($mcount) mailboxes present."
						$Summary += "Send as Permission explicity configured on $($s) of these mailboxes. Full Access Permission explicity configured on $($f) of these mailboxes. Send on Behalf explicity configured on $($b) of these mailboxes and $($n) mailbox has inherited permissions only.<td></table><br>"
					}
					$Header = "
       <body>
       <font size=""1"" face=""Arial,sans-serif"">
       <h3 align=""center"">Mailbox Send As, Full Permission and Send on Behalf Report</h3>
       <h4 align=""center"">Generated $((Get-Date).ToString())</h4>"
					$inh += "</tr></table><br>"
					$gen += "</tr></table><br>"
					$Footer = "</table></center><br><br>
       Scripted by <a href=""http://www.get-mailbox.org"">Serkan Varoglu</a>.  
       Elapsed Time To Complete This Report: $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString())</body></html>"
					if (($dbnull -eq 0) -and ($mbnull -eq 0))
					{
						$Output = $Header + $Summary + $gen + $inh + "<br><hr /><br>" + $Table + $Footer
					}
					elseif (($dbnull -ne 0) -and ($mbnull -eq 0))
					{
						$Output = $Header + $Summary + $gen + $inh + "<br><hr /><br>" + $Table + $Footer
					}
					else
					{
						if (($s -eq 0) -and ($f -eq 0) -and ($b -eq 0))
						{
							$Note = "<center></font><b>Mailbox for $($Mailbox.name) ( $($Mailbox.primarysmtpaddress) ), does not have any explicit permissions set for Send As, Full Access or Send on Behalf</b></center>"
						}
						$Output = $Header + $Note + $Table + $Footer
					}
				}
				else
				{
					$Header = "
       <body>
       <font size=""1"" face=""Arial,sans-serif"">
       <h3 align=""center"">Mailbox Send As, Full Permission and Send on Behalf Report</h3>
       <a name=""top""><h4 align=""center"">Generated $((Get-Date).ToString())</h4></a>
       "
					$inh += "</tr></table><br>"
					$gen += "</tr></table><br>"
					$Footer = "</table></center><br><br>
       <font size=""1"" face=""Arial,sans-serif"">Scripted by <a href=""http://www.get-mailbox.org"">Serkan Varoglu</a>.  
       Elapsed Time To Complete This Report: $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString())</font></body></html>"
					$Output = $Header + $gen + $Table + $Footer
					
				}
				$Output | Out-File $HTMLReport
				
				$Reboot = $false
				
			}
			#endregion
			#region Option 70) Export Office 365 User Last Logon Date to CSV File
			70 {
				if ($O365Creds -eq $null)
				{
					$O365Creds = Get-Credential -Message "Enter your 0365 credentials!"
				}
				O365ExportLastLogonDate
				$Reboot = $false
				
			}
			#endregion
			#region Option 71) List all Distribution Groups and their Membership in Office 365
			71 {
				if ($O365Creds -eq $null)
				{
					$O365Creds = Get-Credential -Message "Enter your 0365 credentials!"
				}
				O365ListDistGroupsAndMemberships -O365Creds $O365Creds
			}
			#endregion
			#region Option 72) Office 365 Mail Traffic Statistics by User
			72 {
				if ($O365Creds -eq $null)
				{
					$O365Creds = Get-Credential -Message "Enter your 0365 credentials!"
				}
				O365MailTrafficStatsbyUser -O365Creds $O365Creds
			}
			#endregion
			#region Option 73) Export a Licence reconciliation report from Office 365
			73 {
				if ($O365Creds -eq $null)
				{
					$O365Creds = Get-Credential -Message "Enter your 0365 credentials!"
				}
				O365ExportLicenseReconcilation ($O365Creds)
			}
			#endregion
			#region Option 74) Export mailbox permissions from Office 365 to CSV file
			74 {
				#region user input
				$HTMLReport = Read-Host "Specifiy alternate path and name for report file. Default is ""mailboxpermissionsreport.html"""
				if ($HTMLReport = "")
				{
					$HTMLReport = "O365MailboxFolderPermissionsReport.html"
				}
				$ShowInheritedYN = Read-Host "List inherited SendAs and Full Access permissions?[Y/N] Default is [Y]"
				switch ($ShowInheritedYN)
				{
					Y{ $ShowInherited = $true }
					N{ $ShowInherited = $false }
					default { $ShowInherited = $true }
				}
				$ShowSelfYN = Read-Host "List NT Authority\Self Permission ?[Y/N] Default is [Y]"
				switch ($ShowInheritedYN)
				{
					Y{ $ShowSelf = $true }
					N{ $ShowSelf = $false }
					default { $ShowSelf = $true }
				}
				$MailboxYN = Read-Host "Specify a single mailbox only?[Y/N] Default is [N]"
				switch ($MailboxYN)
				{
					Y{ $Mailbox = Read-Host "Enter mailbox name" }
					N{ $Mailbox = $null }
					default { $Mailbox = $null }
				}
				if ($Mailbox -eq $null)
				{
					$dbYN = Read-Host "Specify a single database only?[Y/N] Default is [N]"
					switch ($dbYN)
					{
						Y{ $Database = Read-Host "Enter database name" }
						N{ $Database = $null }
						default { $Database = $null }
					}
				}
				if ($O365Creds -eq $null)
				{
					$O365Creds = Get-Credential -Message "Enter your 0365 credentials!"
				}
				#endregion
				
				ExportMBXFolderPermissions -HTMLReport $HTMLReport -ShowInherited:$ShowInherited -ShowSelf:$ShowSelf -Database $Database -Mailbox $Mailbox -O365 -O365Creds $O365Creds 
			}
			#endregion
			#region Option 75) Microsoft 365 Mailboxes with Synchronized Mobile Devices - by Tony Redmond
			75 {
				# An example script to show how to extract mobile device statistics from devices registred with Exchange Online mailboxes
                # https://github.com/12Knocksinna/Office365itpros/blob/master/Report-MobileDevices.PS1

        $directory23 = "C:\mdm\"

if (-not (Test-Path -Path $directory23 -PathType Container)) {
    New-Item -Path $directory23 -ItemType Directory
}
        
        $HtmlHead ="<html>
	    <style>
	    BODY{font-family: Arial; font-size: 8pt;}
	    H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	    H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	    H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	    TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	    TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	    TD{border: 1px solid #969595; padding: 5px; }
	    td.pass{background: #B7EB83;}
	    td.warn{background: #FFF275;}
	    td.fail{background: #FF2626; color: #ffffff;}
	    td.info{background: #85D4FF;}
	    </style>
	    <body>
           <div align=center>
           <p><h1>Microsoft 365 Mailboxes with Synchronized Mobile Devices</h1></p>
           <p><h3>Generated: " + (Get-Date -format 'dd-MMM-yyyy hh:mm tt') + "</h3></p></div>"

$Version = "1.0"
$HtmlReportFile = "C:\mdm\MobileDevices.html"
$CSVReportFile = "C:\mdm\MobileDevices.csv"

Connect-ExchangeOnline

$Organization = Get-OrganizationConfig | Select-Object -ExpandProperty DisplayName
[array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Sort-Object DisplayName
If (!($Mbx)) { Write-Host "Unable to find any user mailboxes..." ; break }

$Report = [System.Collections.Generic.List[Object]]::new() 

[int]$i = 0
ForEach ($M in $Mbx) {
 $i++
 Write-Host ("Scanning mailbox {0} for registered mobile devices... {1}/{2}" -f $M.DisplayName, $i, $Mbx.count)
 [array]$Devices = Get-MobileDevice -Mailbox $M.DistinguishedName
 ForEach ($Device in $Devices) {
   $DaysSinceLastSync = $Null; $DaySinceFirstSync = $Null; $SyncStatus = "OK"
   $DeviceStats = Get-ExoMobileDeviceStatistics -Identity $Device.DistinguishedName
   If ($Device.FirstSyncTime) {
      $DaysSinceFirstSync = (New-TimeSpan $Device.FirstSyncTime).Days }
   If (!([string]::IsNullOrWhiteSpace($DeviceStats.LastSuccessSync))) {
      $DaysSinceLastSync = (New-TimeSpan $DeviceStats.LastSuccessSync).Days }
   If ($DaysSinceLastSync -gt 30)  {
      $SyncStatus = ("Warning: {0} days since last sync" -f $DaysSinceLastSync) }
   If ($Null -eq $DaysSinceLastSync) {
      $SyncStatus = "Never synched" 
      $DeviceStatus = "Unknown" 
   } Else {
      $DeviceStatus =  $DeviceStats.Status }
   $ReportLine = [PSCustomObject]@{
     DeviceId            = $Device.DeviceId
     DeviceOS           = $Device.DeviceOS
     Model              = $Device.DeviceModel
     UA                 = $Device.DeviceUserAgent
     User               = $Device.UserDisplayName
     UPN                = $M.UserPrincipalName
     FirstSync          = $Device.FirstSyncTime
     DaysSinceFirstSync = $DaysSinceFirstSync
     LastSync           = $DeviceStats.LastSuccessSync
     DaysSinceLastSync  = $DaysSinceLastSync
     SyncStatus         = $SyncStatus
     Status             = $DeviceStatus
     Policy             = $DeviceStats.DevicePolicyApplied
     State              = $DeviceStats.DeviceAccessState
     LastPolicy         = $DeviceStats.LastPolicyUpdateTime
     DeviceDN           = $Device.DistinguishedName }
   $Report.Add($ReportLine)
 } #End Devices
} #End Mailboxes
[array]$SyncMailboxes = $Report | Sort-Object UPN -Unique | Select-Object UPN
[array]$SyncDevices = $Report | Sort-Object DeviceId -Unique | Select-Object DeviceId
[array]$SyncDevices30 = $Report | Where-Object {$_.DaysSinceLastSync -gt 30} 
$HtmlReport = $Report | Select-Object DeviceId, DeviceOS, Model, UA, User, UPN, FirstSync, DaysSinceFirstSync, LastSync, DaysSinceLastSync | Sort-Object UPN | ConvertTo-Html -Fragment

# Create the HTML report
$Htmltail = "<p>Report created for: " + ($Organization) + "</p><p>" +
             "<p>Number of mailboxes:                          " + $Mbx.count + "</p>" +
             "<p>Number of users synchronzing devices:         " + $SyncMailboxes.count + "</p>" +
             "<p>Number of synchronized devices:               " + $SyncDevices.count + "</p>" +
             "<p>Number of devices not synced in last 30 days: " + $SyncDevices30.count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------" +
             "<p>Microsoft 365 Mailboxes with Synchronized Mobile Devices<b>" + $Version + "</b>"	
$HtmlReport = $HtmlHead + $HtmlReport + $HtmlTail
$HtmlReport | Out-File $HtmlReportFile  -Encoding UTF8

Write-Host ""
Write-Host "All done"
Write-Host ""
Write-Host ("{0} Mailboxes with synchronized devices" -f $SyncMailboxes.count)
Write-Host ("{0} Individual devices found" -f $SyncDevices.count)

$Report | Export-CSV -NoTypeInformation $CSVReportFile
Write-Host ("Output files are available in {0} and {1}" -f $HtmlReportFile, $CSVReportFile)
Start-Sleep -s 4

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
			}
			#endregion
			#region Option 98) Exit and restart
			98 {
				#      Exit and restart
				Stop-Transcript
				restart-computer -computername localhost -force
			}
			#endregion
			#region Option 99) Exit
			99 {
				#      Exit
				if (($WasInstalled -eq $false) -and (Get-Module BitsTransfer))
				{
					Write-Host "BitsTransfer: Removing..." -NoNewLine
					Remove-Module BitsTransfer
					Write-Host "`b`b`b`b`b`b`b`b`b`b`bremoved!   " -ForegroundColor Green
				}
				popd
				Write-Host "Exiting..."
				Stop-Transcript
			}
			#endregion                 
			default { Write-Host "You haven't selected any of the available options. " }
		}
	}
	while ($Choice -ne 99)
}
#region ----- Menu 2016 -----
######################################################
#    This section is for the Windows 2016 OS         #
######################################################

function Code2016
{
	
	# Start code block for Windows 2016 Server
	
	$Menu2016 = {
		
		write-host " ********************************************************************" -ForegroundColor Cyan
		write-host " Exchange Server 2016 on Windows Server 2016" -ForegroundColor Cyan
		Write-Host "        --- keep it simple, but significant ---" -ForegroundColor Gray
		Write-Host " >>> MSB365 2018 Suite - www.msb365.blog <<<" -ForegroundColor Cyan
		write-host " ********************************************************************" -ForegroundColor Cyan
		write-host " "
		write-host " Please select an option from the list below:" -ForegroundColor White
		write-host " "
		write-host " EXCHANGE SETUP PREREQUISITES (* Exchange media required!)" -ForegroundColor Yellow
		write-host " ---------------------------------------------------------" -foregroundcolor yellow
		write-host " "
		write-host " 1) Launch Windows Update" -ForegroundColor White
		write-host " 2) Check Prerequisites for Mailbox role | Multirole" -ForegroundColor White
		write-host " 3) Check Prerequisites for Edge role" -ForegroundColor White
		write-host " "
		write-host " 4) Install Mailbox prerequisites - Part 1 - CU3+" -ForegroundColor white
		write-host " 5) Install Mailbox prerequisites - Part 2 - CU3+" -ForegroundColor white
		write-host " 6) Install Edge Transport Server prerequisites - CU3 +" -ForegroundColor white
		write-host " "
		write-host " 7) Install - One-Off - Windows Features [MBX]" -ForegroundColor white
		write-host " 8) Install - One Off - Unified Communications Managed API 4.0" -ForegroundColor white
		write-host " "
		write-host " 9) Prepare Schema *" -ForegroundColor White
		write-host " 10) Prepare Active Directory and Domains *" -ForegroundColor White
		write-host " "
		write-host " 11) Set Power Plan to High Performance" -ForegroundColor white
		write-host " 12) Disable Power Management for NICs." -ForegroundColor white
		write-host " 13) Disable SSL 3.0 Support" -ForegroundColor white
		write-host " 14) Disable RC4 Support" -ForegroundColor white
		write-host " "
		write-host " "
		Write-Host " EXCHANGE SETUP TASKS (* Exchange media required!) " -ForegroundColor yellow
		write-host " ---------------------------------------------------------" -foregroundcolor yellow
		write-host " "
		Write-Host " 30) Start Exchange Server setup *" -ForegroundColor Magenta
		write-host " "
		write-host " "
		write-host " POST EXCHANGE 2016 INSTALL" -ForegroundColor Yellow
		write-host " ---------------------------------------------------------" -foregroundcolor yellow
		write-host " "
		write-host " 40) Configure PageFile to RAM + 10 MB" -ForegroundColor green
		Write-Host " 41) Show Exchange URI" -ForegroundColor white
		Write-Host " 42) Configure Exchange URLs" -ForegroundColor white
		Write-Host " 43) Disable UAC" -ForegroundColor white
		Write-Host " 44) Disable Windows Firewall" -ForegroundColor white
		write-host " "
		Write-Host " 45) Create receive connector" -ForegroundColor White
		write-host " 46) Create send connector" -ForegroundColor White
		write-host " 47) Create DAG" -ForegroundColor white
		#write-host " 48) -Create Exchange Hybrid mode" -ForegroundColor white
		write-host " "
		write-host " 49) Create Certificate request" -ForegroundColor White
		Write-Host " 50) set mailaddress policies" -ForegroundColor White
		write-host " "
		write-host " 51) Enable UM for all Mailboxes" -ForegroundColor White
		write-host " 52) Remove  old EAS devices" -ForegroundColor White
		#write-host " 53) -Deploy Microsoft Teams Desktop Client" -ForegroundColor White
		write-host " "
		#write-host " 54) -Order certificate >>GO DADDY<<" -ForegroundColor White
		#write-host " 55) -Order certificate >>DIGICERT<<" -ForegroundColor White
		write-host " "
		write-host " "
		write-host " OPERATING EXCHANGE" -ForegroundColor Yellow
		write-host " ---------------------------------------------------------" -foregroundcolor yellow
		write-host " "
		write-host " 60) Generate Health Report for an Exchange Server 2016/2013/2010 Environment" -ForegroundColor white
		write-host " 61) Generate Exchange Environment Reports" -ForegroundColor white
		#write-host " 62) -Generate Mailbox Size and Information Reports" -ForegroundColor white
		write-host " 63) Generate Reports for Exchange ActiveSync Device Statistics" -ForegroundColor white
		#write-host " 64) -Exchange Analyzer" -ForegroundColor white
		write-host " 65) Generate Report Total Emails Sent and Received Per Day and Size" -ForegroundColor white
		write-host " 66) Generate HTML Report for Mailbox Permissions" -ForegroundColor white
		write-host " "
		write-host " "
		write-host " Office 365 Operation" -ForegroundColor Yellow
		write-host " ---------------------------------------------------------" -foregroundcolor yellow
		write-host " "
		write-host " 70) Export Office 365 User Last Logon Date to CSV File" -ForegroundColor white
		write-host " 71) List all Distribution Groups and their Membership in Office 365" -ForegroundColor white
		write-host " 72) Office 365 Mail Traffic Statistics by User" -ForegroundColor white
		write-host " 73) Export a Licence reconciliation report from Office 365" -ForegroundColor white
		write-host " 74) Export mailbox permissions from Office 365 to CSV file" -ForegroundColor white
		write-host " 75) Microsoft 365 Mailboxes with Synchronized Mobile Devices" -ForegroundColor white
		write-host " "
		write-host " "
		write-host " OPERATING EXCHANGE" -ForegroundColor Yellow
		write-host " ---------------------------------------------------------" -foregroundcolor yellow
		write-host " "
		write-host " 98) Restart the Server" -foregroundcolor red
		write-host " 99) Exit" -foregroundcolor cyan
		write-host " "
		write-host " Select an option.. [1-99]? " -foregroundcolor white -nonewline
	}
	#endregion   
	#region Check Mailox Requirements      
	# Check Mailox Requirements
	function check-MBXprereq
	{
		write-host " "
		write-host "Checking all requirements for the Mailbox Role in Exchange Server 2016 on Windows Server 2016....." -foregroundcolor yellow
		write-host " "
		start-sleep 2
		#endregion   
		#region .NET Check - Removed as Windows 2016 has .NET 4.6.2 by default         
		# .NET Check - Removed as Windows 2016 has .NET 4.6.2 by default
		
		# Windows Feature Check
		$values = @("NET-Framework-45-Features", "RPC-over-HTTP-proxy", "RSAT-Clustering", "RSAT-Clustering-CmdInterface", "RSAT-Clustering-Mgmt", "RSAT-Clustering-PowerShell", "Web-Mgmt-Console", "WAS-Process-Model", "Web-Asp-Net45", "Web-Basic-Auth", "Web-Client-Auth", "Web-Digest-Auth", "Web-Dir-Browsing", "Web-Dyn-Compression", "Web-Http-Errors", "Web-Http-Logging", "Web-Http-Redirect", "Web-Http-Tracing", "Web-ISAPI-Ext", "Web-ISAPI-Filter", "Web-Lgcy-Mgmt-Console", "Web-Metabase", "Web-Mgmt-Console", "Web-Mgmt-Service", "Web-Net-Ext45", "Web-Request-Monitor", "Web-Server", "Web-Stat-Compression", "Web-Static-Content", "Web-Windows-Auth", "Web-WMI", "Windows-Identity-Foundation")
		foreach ($item in $values)
		{
			$val = get-Windowsfeature $item
			If ($val.installed -eq $true)
			{
				write-host "The Windows Feature"$item" is " -nonewline
				write-host "installed." -ForegroundColor green
			}
			else
			{
				write-host "The Windows Feature"$item" is " -nonewline
				write-host "not installed!" -ForegroundColor red
			}
		}
		#endregion   
		
		#region Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit         
		# Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit 
		$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
		if ($val.DisplayVersion -ne "5.0.8308.0")
		{
			if ($val.DisplayVersion -ne "5.0.8132.0")
			{
				if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A41CBE7D-949C-41DD-9869-ABBD99D753DA}") -eq $false)
				{
					write-host "No version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
					write-host "not installed!" -ForegroundColor red
					write-host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992."
				}
				else
				{
					write-host "The Preview version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
					write-host "installed." -ForegroundColor red
					write-host "This is the incorrect version of UCMA. " -nonewline -ForegroundColor red
					write-host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992."
				}
			}
			else
			{
				write-host "The wrong version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
				write-host "installed." -ForegroundColor red
				write-host "This is the incorrect version of UCMA. " -nonewline -ForegroundColor red
				write-host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992."
			}
		}
		else
		{
			write-host "The correct version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
			write-host "installed." -ForegroundColor green
		}
	} # End function check-MBXprereq
	#endregion   
	
	#region Check Edge Requirements  
	# Check Edge Requirements
	function Check-EdgePrereq
	{
		
		write-host " "
		write-host "Checking all requirements for the Edge Transport Role in Exchange Server 2016 on Windows Server 2016....." -foregroundcolor yellow
		write-host " "
		start-sleep 2
		#endregion   
		
		#region Check .NET version - Removed as Windows 2016 has .NET 4.6.2 by default       
		# Check .NET version - Removed as Windows 2016 has .NET 4.6.2 by default
		
		# Windows Feature AD LightWeight Services
		$values = @("ADLDS")
		foreach ($item in $values)
		{
			$val = get-Windowsfeature $item
			If ($val.installed -eq $true)
			{
				write-host "The Windows Feature"$item" is " -nonewline
				write-host "installed." -ForegroundColor green
				write-host " "
			}
			else
			{
				write-host "The Windows Feature"$item" is " -nonewline
				write-host "not installed!" -ForegroundColor red
				write-host " "
			}
		}
	}
	#endregion   
	
	#region Start Windows Defender function 
	# Start Windows Defender function
	function WindowsDefender
	{
		write-host " "
		write-host "Windows Defender exclusions:" -ForegroundColor cyan
		if (Get-Module Defender -ListAvailable)
		{
			try
			{
				# Noderunner exclusion
				$ExchangeProcess = "$exinstall\Bin\Search\Ceres\Runtime\1.0\Noderunner.exe"
				Add-MpPreference -ExclusionProcess $ExchangeProcess
				write-host "Added " -foregroundcolor white -nonewline
				write-host "Process exclusions" -foregroundcolor green -nonewline
				write-host " successfully!" -foregroundcolor white
			}
			catch
			{
				Write-Warning $_.Exception.Message
			}
			try
			{
				# System Drive
				$Drive = $env:SystemDrive
				$ExchangeSetupLog = "$drive\ExchangeSetupLogs\ExchangeSetup.log"
				Add-MpPreference -ExclusionPath $ExchangeSetupLog
				write-host "Added " -foregroundcolor white -nonewline
				write-host "Setup Log exclusions" -foregroundcolor green -nonewline
				write-host " successfully!" -foregroundcolor white
			}
			catch
			{
				Write-Warning $_.Exception.Message
			}
			try
			{
				# Exchange Installation Director
				Add-MpPreference -ExclusionPath $exinstall
				write-host "Added " -foregroundcolor white -nonewline
				write-host "Exchange Install directory exclusions" -foregroundcolor green -nonewline
				write-host " successfully!" -foregroundcolor white
			}
			catch
			{
				Write-Warning $_.Exception.Message
			}
		}
		else
		{
			Write-Warning "Windows Defender PowerShell module not available."
		}
		write-host " "
	} # End Windows Defender function
	
	Do
	{
		if ($Reboot -eq $true) { Write-Host "`t`t`t`t`t`t`t`t`t`n`t`t`t`tREBOOT REQUIRED!`t`t`t`n`t`t`t`t`t`t`t`t`t`n`t`tDO NOT INSTALL EXCHANGE BEFORE REBOOTING!`t`t`n`t`t`t`t`t`t`t`t`t" -backgroundcolor red -foregroundcolor black }
		if ($Choice -ne "None") { Write-Host "Last command: "$Choice -foregroundcolor Yellow }
		invoke-command -scriptblock $Menu2016
		$Choice = Read-Host
		#endregion   
		
		#region Functions
		Function ShowEXCURI
		{
			param
			(
				[parameter(Mandatory = $false)]
				[String[]]
				$server = (Get-ExchangeServer).fqdn
			)
			
			#get all EXC server				
			#$server = (Get-ExchangeServer).fqdn
			#...................................
			# Script
			#...................................
			
			Begin
			{
				
				#Add Exchange snapin if not already loaded in the PowerShell session
				if (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1)
				{
					. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
					Connect-ExchangeServer -auto -AllowClobber
				}
				else
				{
					Write-Warning "Exchange Server management tools are not installed on this computer."
					EXIT
				}
			}
			
			Process
			{
				
				foreach ($i in $server)
				{
					if ((Get-ExchangeServer $i -ErrorAction SilentlyContinue).IsClientAccessServer)
					{
						Write-Host "----------------------------------------"
						Write-Host " Querying $i"
						Write-Host "----------------------------------------`r`n"
						Write-Host "`r`n"
						
						$OA = Get-OutlookAnywhere -Server $i -AdPropertiesOnly | Select InternalHostName, ExternalHostName
						Write-Host "Outlook Anywhere"
						Write-Host " - Internal: $($OA.InternalHostName)"
						Write-Host " - External: $($OA.ExternalHostName)"
						Write-Host "`r`n"
						
						$OWA = Get-OWAVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalURL, ExternalURL
						Write-Host "Outlook Web App"
						Write-Host " - Internal: $($OWA.InternalURL)"
						Write-Host " - External: $($OWA.ExternalURL)"
						Write-Host "`r`n"
						
						$ECP = Get-ECPVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalURL, ExternalURL
						Write-Host "Exchange Control Panel"
						Write-Host " - Internal: $($ECP.InternalURL)"
						Write-Host " - External: $($ECP.ExternalURL)"
						Write-Host "`r`n"
						
						$OAB = Get-OABVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalURL, ExternalURL
						Write-Host "Offline Address Book"
						Write-Host " - Internal: $($OAB.InternalURL)"
						Write-Host " - External: $($OAB.ExternalURL)"
						Write-Host "`r`n"
						
						$EWS = Get-WebServicesVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalURL, ExternalURL
						Write-Host "Exchange Web Services"
						Write-Host " - Internal: $($EWS.InternalURL)"
						Write-Host " - External: $($EWS.ExternalURL)"
						Write-Host "`r`n"
						
						$MAPI = Get-MAPIVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalURL, ExternalURL
						Write-Host "MAPI"
						Write-Host " - Internal: $($MAPI.InternalURL)"
						Write-Host " - External: $($MAPI.ExternalURL)"
						Write-Host "`r`n"
						
						$EAS = Get-ActiveSyncVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalURL, ExternalURL
						Write-Host "ActiveSync"
						Write-Host " - Internal: $($EAS.InternalURL)"
						Write-Host " - External: $($EAS.ExternalURL)"
						Write-Host "`r`n"
						
						$AutoD = Get-ClientAccessServer $i | Select AutoDiscoverServiceInternalUri
						Write-Host "Autodiscover"
						Write-Host " - Internal SCP: $($AutoD.AutoDiscoverServiceInternalUri)"
						Write-Host "`r`n"
						
					}
					else
					{
						Write-Host -ForegroundColor Yellow "$i is not a Client Access server."
					}
				}
			}
			End
			{
				
				Write-Host "Finished querying all servers specified."
				
			}
		}
		
		function ConfigureEXCURL($server, $InternalURL, $ExternalURL, $AutodiscoverSCP, $InternalSSL, $ExternalSSL)
		{
			Begin
			{
				
				#Add Exchange snapin if not already loaded in the PowerShell session
				if (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1)
				{
					. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
					Connect-ExchangeServer -auto -AllowClobber
				}
				else
				{
					Write-Warning "Exchange Server management tools are not installed on this computer."
					Return
				}
			}
			
			Process
			{
				
				foreach ($i in $server)
				{
					if ((Get-ExchangeServer $i -ErrorAction SilentlyContinue).IsClientAccessServer)
					{
						Write-Host "----------------------------------------"
						Write-Host " Configuring $i"
						Write-Host "----------------------------------------`r`n"
						Write-Host "Values:"
						Write-Host " - Internal URL: $InternalURL"
						Write-Host " - External URL: $ExternalURL"
						Write-Host " - Outlook Anywhere internal SSL required: $InternalSSL"
						Write-Host " - Outlook Anywhere external SSL required: $ExternalSSL"
						Write-Host "`r`n"
						
						Write-Host "Configuring Outlook Anywhere URLs"
						$OutlookAnywhere = Get-OutlookAnywhere -Server $i
						$OutlookAnywhere | Set-OutlookAnywhere -ExternalHostname $externalurl -InternalHostname $internalurl `
															   -ExternalClientsRequireSsl $ExternalSSL -InternalClientsRequireSsl $InternalSSL `
															   -ExternalClientAuthenticationMethod $OutlookAnywhere.ExternalClientAuthenticationMethod
						
						if ($externalurl -eq "")
						{
							Write-Host "Configuring Outlook Web App URLs"
							Get-OwaVirtualDirectory -Server $i | Set-OwaVirtualDirectory -ExternalUrl $null -InternalUrl https://$internalurl/owa
							
							Write-Host "Configuring Exchange Control Panel URLs"
							Get-EcpVirtualDirectory -Server $i | Set-EcpVirtualDirectory -ExternalUrl $null -InternalUrl https://$internalurl/ecp
							
							Write-Host "Configuring ActiveSync URLs"
							Get-ActiveSyncVirtualDirectory -Server $i | Set-ActiveSyncVirtualDirectory -ExternalUrl $null -InternalUrl https://$internalurl/Microsoft-Server-ActiveSync
							
							Write-Host "Configuring Exchange Web Services URLs"
							Get-WebServicesVirtualDirectory -Server $i | Set-WebServicesVirtualDirectory -ExternalUrl $null -InternalUrl https://$internalurl/EWS/Exchange.asmx
							
							Write-Host "Configuring Offline Address Book URLs"
							Get-OabVirtualDirectory -Server $i | Set-OabVirtualDirectory -ExternalUrl $null -InternalUrl https://$internalurl/OAB
							
							Write-Host "Configuring MAPI/HTTP URLs"
							Get-MapiVirtualDirectory -Server $i | Set-MapiVirtualDirectory -ExternalUrl $null -InternalUrl https://$internalurl/mapi
						}
						else
						{
							Write-Host "Configuring Outlook Web App URLs"
							Get-OwaVirtualDirectory -Server $i | Set-OwaVirtualDirectory -ExternalUrl https://$externalurl/owa -InternalUrl https://$internalurl/owa
							
							Write-Host "Configuring Exchange Control Panel URLs"
							Get-EcpVirtualDirectory -Server $i | Set-EcpVirtualDirectory -ExternalUrl https://$externalurl/ecp -InternalUrl https://$internalurl/ecp
							
							Write-Host "Configuring ActiveSync URLs"
							Get-ActiveSyncVirtualDirectory -Server $i | Set-ActiveSyncVirtualDirectory -ExternalUrl https://$externalurl/Microsoft-Server-ActiveSync -InternalUrl https://$internalurl/Microsoft-Server-ActiveSync
							
							Write-Host "Configuring Exchange Web Services URLs"
							Get-WebServicesVirtualDirectory -Server $i | Set-WebServicesVirtualDirectory -ExternalUrl https://$externalurl/EWS/Exchange.asmx -InternalUrl https://$internalurl/EWS/Exchange.asmx
							
							Write-Host "Configuring Offline Address Book URLs"
							Get-OabVirtualDirectory -Server $i | Set-OabVirtualDirectory -ExternalUrl https://$externalurl/OAB -InternalUrl https://$internalurl/OAB
							
							Write-Host "Configuring MAPI/HTTP URLs"
							Get-MapiVirtualDirectory -Server $i | Set-MapiVirtualDirectory -ExternalUrl https://$externalurl/mapi -InternalUrl https://$internalurl/mapi
						}
						
						Write-Host "Configuring Autodiscover"
						if ($AutodiscoverSCP -ne "")
						{
							Get-ClientAccessServer $i | Set-ClientAccessServer -AutoDiscoverServiceInternalUri https://$AutodiscoverSCP/Autodiscover/Autodiscover.xml
						}
						else
						{
							Get-ClientAccessServer $i | Set-ClientAccessServer -AutoDiscoverServiceInternalUri https://$internalurl/Autodiscover/Autodiscover.xml
						}
						
						
						Write-Host "`r`n"
					}
					else
					{
						Write-Host -ForegroundColor Yellow "$i is not a Client Access server."
					}
				}
			}
			End
			{
				
				Write-Host "Finished processing all servers specified. Consider running Get-CASHealthCheck.ps1 to test your Client Access namespace and SSL configuration."
				Write-Host "Refer to http://exchangeserverpro.com/testing-exchange-server-2013-client-access-server-health-with-powershell/ for more details."
				
			}
		}
		
		function generateHealthReport
		{
			#      Generate Health Report for an Exchange Server 2016/2013/2010 Environment
                           <#
.SYNOPSIS
Test-ExchangeServerHealth.ps1 - Exchange Server Health Check Script.

.DESCRIPTION 
Performs a series of health checks on Exchange servers and DAGs
and outputs the results to screen, and optionally to log file, HTML report,
and HTML email.

Use the ignorelist.txt file to specify any servers, DAGs, or databases you
want the script to ignore (eg test/dev servers).

.OUTPUTS
Results are output to screen, as well as optional log file, HTML report, and HTML email

.PARAMETER Server
Perform a health check of a single server

.PARAMETER ReportMode
Set to $true to generate a HTML report. A default file name is used if none is specified.

.PARAMETER ReportFile
Allows you to specify a different HTML report file name than the default.

.PARAMETER SendEmail
Sends the HTML report via email using the SMTP configuration within the script.

.PARAMETER AlertsOnly
Only sends the email report if at least one error or warning was detected.

.PARAMETER Log
Writes a log file to help with troubleshooting.

.EXAMPLE
.\Test-ExchangeServerHealth.ps1
Checks all servers in the organization and outputs the results to the shell window.

.EXAMPLE
.\Test-ExchangeServerHealth.ps1 -Server HO-EX2010-MB1
Checks the server HO-EX2010-MB1 and outputs the results to the shell window.

.EXAMPLE
.\Test-ExchangeServerHealth.ps1 -ReportMode -SendEmail
Checks all servers in the organization, outputs the results to the shell window, a HTML report, and
emails the HTML report to the address configured in the script.

.LINK
https://practical365.com/exchange-server/powershell-script-exchange-server-health-check-report/

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:   http://paulcunningham.me
* Twitter:   https://twitter.com/paulcunningham
* LinkedIn:  http://au.linkedin.com/in/cunninghamp/
* Github:    https://github.com/cunninghamp

For more Exchange Server tips, tricks and news
check out Exchange Server Pro.

* Website:   https://practical365.com
* Twitter:   https://twitter.com/practical365

Additional Credits (code contributions and testing):
- Chris Brown, http://twitter.com/chrisbrownie
- Ingmar Brückner
- John A. Eppright
- Jonas Borelius
- Thomas Helmdach
- Bruce McKay
- Tony Holdgate
- Ryan
- Rob Silver
- andrewcr7, https://github.com/andrewcr7

License:

The MIT License (MIT)

Copyright (c) 2017 Paul Cunningham

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Change Log
V1.00, 05/07/2012 - Initial version
V1.01, 05/08/2012 - Minor bug fixes and removed Edge Tranport checks
V1.02, 05/05/2013 - A lot of bug fixes, updated SMTP to use Send-MailMessage, added DAG health check.
V1.03, 04/08/2013 - Minor bug fixes
V1.04, 19/08/2013 - Added Exchange 2013 compatibility, added option to output a log file, converted many
                    sections of code to use pre-defined strings, fixed -AlertsOnly parameter, improved summary 
                    sections of report to be more readable and include DAG summary
V1.05, 23/08/2013 - Added workaround for Test-ServiceHealth error for Exchange 2013 CAS-only servers
V1.06, 28/10/2013 - Added workaround for Test-Mailflow error for Exchange 2013 Mailbox servers.
                  - Added workaround for Exchange 2013 mail test.
                             - Added localization strings for service health check errors for non-English systems.
                             - Fixed an uptime calculation bug for some regional settings.
                             - Excluded recovery databases from active database calculation.
                             - Fixed bug where high transport queues would not count as an alert.
                  - Fixed error thrown when Site attribute can't be found for Exchange 2003 servers.
                  - Fixed bug causing Exchange 2003 servers to be added to the report twice.
V1.07, 24/11/2013 - Fixed bug where disabled content indexes were counted as failed.
V1.08, 29/06/2014 - Fixed bug with DAG reporting in mixed Exchange 2010/2013 orgs.
V1.09, 06/07/2014 - Fixed bug with DAG member replication health reporting for mixed Exchange 2010/2013 orgs.
V1.10, 19/08/2014 - Fixed bug with E14 replication health not testing correct server.
V1.11, 11/02/2015 - Added queue length to Transport queue result in report.
V1.12, 05/03/2015 - Fixed bug with color-coding in report for Transport Queue length.
V1.13, 07/03/2015 - Fixed bug with incorrect function name used sometimes when trying to call Write-LogFile
V1.14, 21/05/2015 - Fixed bug with color-coding in report for Transport Queue length on CAS-only Exchange 2013 servers.
V1.15, 18/11/2015 - Fixed bug with Exchange 2016 version detection.
V1.16, 13/04/2017 - Fixed bugs with recovery DB detection, invalid variables, shadow redundancy queues, and lagged copy detection.
V1.17, 17/05/2017 - Fixed bug with auto-suspended content index detection
#>
			
			
			#Region user input
			$singleorlist = Read-Host "Report for one specific server [1], for a list of servers[2] or for all Exchange servers[3]. Enter for cancel"
			switch ($singleorlist)
			{
				1 { $server = Read-Host "Enter server name" }
				2 { $serverlist = Read-host "Enter path to txt fie with the list of servers" }
				3 { $server = $null; $serverlist = $null }
				default { "No option selected. Exiting"; Return }
			}
			$repmodeYesNo = Read-Host "Generate HTML from report? [Y/N]"
			switch ($repmodeYesNo)
			{
				Y{ $ReportMode = $true }
				N{ $ReportMode = $false }
				default { "No option selected. Exiting"; Return }
				
			}
			if ($ReportMode)
			{
				$reportfilealt = Read-Host "Specifiy alternate path and name for report file. Default is ""exchangeserverhealth.html"" in current path"
				if ($reportfilealt -eq "" -or $reportfilealt -eq $null)
				{
					$ReportFile = "exchangeserverhealth.html"
				}
				else
				{
					$ReportFile = $reportfilealt
				}
				
			}
			$SendMailYesNo = Read-Host "Send e-mail with report? [Y/N] Default is [N]"
			
			switch ($SendMailYesNo)
			{
				Y{ $SendEmail = $true }
				N{ $SendEmail = $false }
				default { $SendEmail = $false }
			}
			if ($SendEmail)
			{
				$AlertsOnlyYN = Read-Host "Send email only if error or warning was detected?[Y/N] Default is [N]"
				switch ($AlertsOnlyYN)
				{
					Y{ $AlertsOnly = $true }
					N{ $AlertsOnly = $false }
					default { $AlertsOnly = $false }
				}
				$smtpServer = Read-Host "Enter SMTP Server"
				$To = Read-Host -Prompt "Enter recipients SMTP address"
				$From = Read-Host -Prompt "Enter senders SMTP address"
				
			}
			$LogYN = Read-Host "Create log file for troubleshooting?[Y/N] Default is [N]"
			switch ($LogYN)
			{
				Y{ $Log = $true; "Log will be saved to $myDir\exchangeserverhealth.log" }
				N{ $Log = $false }
				default { $Log = $false }
			}
			
			#Endregion
			
			
			#...................................
			# Variables
			#...................................
			
			$now = Get-Date #Used for timestamps
			$date = $now.ToShortDateString() #Short date format for email message subject
			[array]$exchangeservers = @() #Array for the Exchange server or servers to check
			[int]$transportqueuehigh = 100 #Change this to set transport queue high threshold. Must be higher than warning threshold.
			[int]$transportqueuewarn = 80 #Change this to set transport queue warning threshold. Must be lower than high threshold.
			$mapitimeout = 10 #Timeout for each MAPI connectivity test, in seconds
			$pass = "Green"
			$warn = "Yellow"
			$fail = "Red"
			$ip = $null
			[array]$serversummary = @() #Summary of issues found during server health checks
			[array]$dagsummary = @() #Summary of issues found during DAG health checks
			[array]$report = @()
			[bool]$alerts = $false
			[array]$dags = @() #Array for DAG health check
			[array]$dagdatabases = @() #Array for DAG databases
			[int]$replqueuewarning = 8 #Threshold to consider a replication queue unhealthy
			$dagreportbody = $null
			
			$myDir = Get-ScriptDirectory
			
			#...................................
			# Modify these Variables (optional)
			#...................................
			
			$reportemailsubject = "Exchange Server Health Report"
			$ignorelistfile = "$myDir\ignorelist.txt"
			$logfile = "$myDir\exchangeserverhealth.log"
			#$ReportFile = "C:\temp\$ReportFile"
			
			#...................................
			# Modify these Email Settings
			#...................................
			
			$smtpsettings = @{
				To	    = "administrator@exchangeserverpro.net"
				From    = "exchangeserver@exchangeserverpro.net"
				Subject = "$reportemailsubject - $now"
				SmtpServer = "smtp.exchangeserverpro.net"
			}
			
			
			#...................................
			# Modify these language 
			# localization strings.
			#...................................
			
			# The server roles must match the role names you see when you run Test-ServiceHealth.
			$casrole = "Client Access Server Role"
			$htrole = "Hub Transport Server Role"
			$mbrole = "Mailbox Server Role"
			$umrole = "Unified Messaging Server Role"
			
			# This should match the word for "Success", or the result of a successful Test-MAPIConnectivity test
			$success = "Success"
			
			#...................................
			# Logfile Strings
			#...................................
			
			$logstring0 = "====================================="
			$logstring1 = " Exchange Server Health Check"
			
			#...................................
			# Initialization Strings
			#...................................
			
			$initstring0 = "Initializing..."
			$initstring1 = "Loading the Exchange Server PowerShell snapin"
			$initstring2 = "The Exchange Server PowerShell snapin did not load."
			$initstring3 = "Setting scope to entire forest"
			
			#...................................
			# Error/Warning Strings
			#...................................
			
			$string0 = "Server is not an Exchange server. "
			$string1 = "Server is not reachable. "
			$string3 = "------ Checking"
			$string4 = "Could not test service health. "
			$string5 = "required services not running. "
			$string6 = "Could not check queue. "
			$string7 = "Public Folder database not mounted. "
			$string8 = "Skipping Edge Transport server. "
			$string9 = "Mailbox databases not mounted. "
			$string10 = "MAPI tests failed. "
			$string11 = "Mail flow test failed. "
			$string12 = "No Exchange Server 2003 checks performed. "
			$string13 = "Server not found in DNS. "
			$string14 = "Sending email. "
			$string15 = "Done."
			$string16 = "------ Finishing"
			$string17 = "Unable to retrieve uptime. "
			$string18 = "Ping failed. "
			$string19 = "No alerts found, and AlertsOnly switch was used. No email sent. "
			$string20 = "You have specified a single server to check"
			$string21 = "Couldn't find the server $server. Script will terminate."
			$string22 = "The file $ignorelistfile could not be found. No servers, DAGs or databases will be ignored."
			$string23 = "You have specified a filename containing a list of servers to check"
			$string24 = "The file $serverlist could not be found. Script will terminate."
			$string25 = "Retrieving server list"
			$string26 = "Removing servers in ignorelist from server list"
			$string27 = "Beginning the server health checks"
			$string28 = "Servers, DAGs and databases to ignore:"
			$string29 = "Servers to check:"
			$string30 = "Checking DNS"
			$string31 = "DNS check passed"
			$string32 = "Checking ping"
			$string33 = "Ping test passed"
			$string34 = "Checking uptime"
			$string35 = "Checking service health"
			$string36 = "Checking Hub Transport Server"
			$string37 = "Checking Mailbox Server"
			$string38 = "Ignore list contains no server names."
			$string39 = "Checking public folder database"
			$string40 = "Public folder database status is"
			$string41 = "Checking mailbox databases"
			$string42 = "Mailbox database status is"
			$string43 = "Offline databases: "
			$string44 = "Checking MAPI connectivity"
			$string45 = "MAPI connectivity status is"
			$string46 = "MAPI failed to: "
			$string47 = "Checking mail flow"
			$string48 = "Mail flow status is"
			$string49 = "No active DBs"
			$string50 = "Finished checking server"
			$string51 = "Skipped"
			$string52 = "Using alternative test for Exchange 2013 CAS-only server"
			$string60 = "Beginning the DAG health checks"
			$string61 = "Could not determine server with active database copy"
			$string62 = "mounted on server that is activation preference"
			$string63 = "unhealthy database copy count is"
			$string64 = "healthy copy/replay queue count is"
			$string65 = "(of"
			$string66 = ")"
			$string67 = "unhealthy content index count is"
			$string68 = "DAGs to check:"
			$string69 = "DAG databases to check"
			
			
			
			#...................................
			# Functions
			#...................................
			
			#This function is used to generate HTML for the DAG member health report
			Function New-DAGMemberHTMLTableCell()
			{
				param ($lineitem)
				
				$htmltablecell = $null
				
				switch ($($line."$lineitem"))
				{
					$null { $htmltablecell = "<td>n/a</td>" }
					"Passed" { $htmltablecell = "<td class=""pass"">$($line."$lineitem")</td>" }
					default { $htmltablecell = "<td class=""warn"">$($line."$lineitem")</td>" }
				}
				
				return $htmltablecell
			}
			
			#This function is used to generate HTML for the server health report
			Function New-ServerHealthHTMLTableCell()
			{
				param ($lineitem)
				
				$htmltablecell = $null
				
				switch ($($reportline."$lineitem"))
				{
					$success { $htmltablecell = "<td class=""pass"">$($reportline."$lineitem")</td>" }
					"Success" { $htmltablecell = "<td class=""pass"">$($reportline."$lineitem")</td>" }
					"Pass" { $htmltablecell = "<td class=""pass"">$($reportline."$lineitem")</td>" }
					"Warn" { $htmltablecell = "<td class=""warn"">$($reportline."$lineitem")</td>" }
					"Access Denied" { $htmltablecell = "<td class=""warn"">$($reportline."$lineitem")</td>" }
					"Fail" { $htmltablecell = "<td class=""fail"">$($reportline."$lineitem")</td>" }
					"Could not test service health. " { $htmltablecell = "<td class=""warn"">$($reportline."$lineitem")</td>" }
					"Unknown" { $htmltablecell = "<td class=""warn"">$($reportline."$lineitem")</td>" }
					default { $htmltablecell = "<td>$($reportline."$lineitem")</td>" }
				}
				
				return $htmltablecell
			}
			
			#This function is used to write the log file if -Log is used
			Function Write-Logfile()
			{
				param ($logentry)
				$timestamp = Get-Date -DisplayHint Time
				"$timestamp $logentry" | Out-File $logfile -Append
			}
			
			#This function is used to test service health for Exchange 2013 CAS-only servers
			Function Test-E15CASServiceHealth()
			{
				param ($e15cas)
				
				$e15casservicehealth = $null
				$servicesrunning = @()
				$servicesnotrunning = @()
				$casservices = @(
					"IISAdmin",
					"W3Svc",
					"WinRM",
					"MSExchangeADTopology",
					"MSExchangeDiagnostics",
					"MSExchangeFrontEndTransport",
					#"MSExchangeHM",
					"MSExchangeIMAP4",
					"MSExchangePOP3",
					"MSExchangeServiceHost",
					"MSExchangeUMCR"
				)
				
				try
				{
					$servicestates = @(Get-WmiObject -ComputerName $e15cas -Class Win32_Service -ErrorAction STOP | Where-Object { $casservices -icontains $_.Name } | Select-Object name, state, startmode)
				}
				catch
				{
					if ($Log) { Write-LogFile $_.Exception.Message }
					Write-Warning $_.Exception.Message
					$e15casservicehealth = "Fail"
				}
				
				if (!($e15casservicehealth))
				{
					$servicesrunning = @($servicestates | Where-Object { $_.StartMode -eq "Auto" -and $_.State -eq "Running" })
					$servicesnotrunning = @($servicestates | Where-Object { $_.Startmode -eq "Auto" -and $_.State -ne "Running" })
					if ($($servicesnotrunning.Count) -gt 0)
					{
						Write-Verbose "Service health check failed"
						Write-Verbose "Services not running:"
						foreach ($service in $servicesnotrunning)
						{
							Write-Verbose "- $($service.Name)"
						}
						$e15casservicehealth = "Fail"
					}
					else
					{
						Write-Verbose "Service health check passed"
						$e15casservicehealth = "Pass"
					}
				}
				return $e15casservicehealth
			}
			
			#This function is used to test mail flow for Exchange 2013 Mailbox servers
			Function Test-E15MailFlow()
			{
				param ($e15mailboxserver)
				
				$e15mailflowresult = $null
				
				Write-Verbose "Creating PSSession for $e15mailboxserver"
				$url = (Get-PowerShellVirtualDirectory -Server $e15mailboxserver -AdPropertiesOnly | Where-Object { $_.Name -eq "Powershell (Default Web Site)" }).InternalURL.AbsoluteUri
				if ($url -eq $null)
				{
					$url = "http://$e15mailboxserver/powershell"
				}
				
				try
				{
					$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $url -ErrorAction STOP
				}
				catch
				{
					Write-Verbose "Something went wrong"
					if ($Log) { Write-LogFile $_.Exception.Message }
					Write-Warning $_.Exception.Message
					$e15mailflowresult = "Fail"
				}
				
				try
				{
					Write-Verbose "Running mail flow test on $e15mailboxserver"
					$result = Invoke-Command -Session $session { Test-Mailflow } -ErrorAction STOP
					$e15mailflowresult = $result.TestMailflowResult
				}
				catch
				{
					Write-Verbose "An error occurred"
					if ($Log) { Write-LogFile $_.Exception.Message }
					Write-Warning $_.Exception.Message
					$e15mailflowresult = "Fail"
				}
				
				Write-Verbose "Mail flow test: $e15mailflowresult"
				Write-Verbose "Removing PSSession"
				Remove-PSSession $session.Id
				
				return $e15mailflowresult
			}
			
			#This function is used to test replication health for Exchange 2010 DAG members in mixed 2010/2013 organizations
			Function Test-E14ReplicationHealth()
			{
				param ($e14mailboxserver)
				
				$e14replicationhealth = $null
				
				#Find an E14 CAS in the same site
				$ADSite = (Get-ExchangeServer $e14mailboxserver).Site
				$e14cas = (Get-ExchangeServer | Where-Object { $_.IsClientAccessServer -and $_.AdminDisplayVersion -match "Version 14" -and $_.Site -eq $ADSite } | Select-Object -first 1).FQDN
				
				Write-Verbose "Creating PSSession for $e14cas"
				$url = (Get-PowerShellVirtualDirectory -Server $e14cas -AdPropertiesOnly | Where-Object { $_.Name -eq "Powershell (Default Web Site)" }).InternalURL.AbsoluteUri
				if ($url -eq $null)
				{
					$url = "http://$e14cas/powershell"
				}
				
				Write-Verbose "Using URL $url"
				
				try
				{
					$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $url -ErrorAction STOP
				}
				catch
				{
					Write-Verbose "Something went wrong"
					if ($Log) { Write-LogFile $_.Exception.Message }
					Write-Warning $_.Exception.Message
					#$e14replicationhealth = "Fail"
				}
				
				try
				{
					Write-Verbose "Running replication health test on $e14mailboxserver"
					#$e14replicationhealth = Invoke-Command -Session $session {Test-ReplicationHealth} -ErrorAction STOP
					$e14replicationhealth = Invoke-Command -Session $session -Args $e14mailboxserver.Name { Test-ReplicationHealth $args[0] } -ErrorAction STOP
				}
				catch
				{
					Write-Verbose "An error occurred"
					if ($Log) { Write-LogFile $_.Exception.Message }
					Write-Warning $_.Exception.Message
					#$e14replicationhealth = "Fail"
				}
				
				#Write-Verbose "Replication health test: $e14replicationhealth"
				Write-Verbose "Removing PSSession"
				Remove-PSSession $session.Id
				
				return $e14replicationhealth
			}
			
			
			#...................................
			# Initialize
			#...................................
			
			#Log file is overwritten each time the script is run to avoid
			#very large log files from growing over time
			if ($Log)
			{
				$timestamp = Get-Date -DisplayHint Time
				"$timestamp $logstring0" | Out-File $logfile
				Write-Logfile $logstring1
				Write-Logfile "  $now"
				Write-Logfile $logstring0
			}
			
			Write-Host $initstring0
			if ($Log) { Write-Logfile $initstring0 }
			
			#Add Exchange 2010 snapin if not already loaded in the PowerShell session
			if (!(Get-PSSnapin | Where-Object { $_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010" }))
			{
				Write-Verbose $initstring1
				if ($Log) { Write-Logfile $initstring1 }
				try
				{
					Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
				}
				catch
				{
					#Snapin was not loaded
					Write-Verbose $initstring2
					if ($Log) { Write-Logfile $initstring2 }
					Write-Warning $_.Exception.Message
					EXIT
				}
				. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
				Connect-ExchangeServer -auto -AllowClobber
			}
			
			
			#Set scope to include entire forest
			Write-Verbose $initstring3
			if ($Log) { Write-Logfile $initstring3 }
			if (!(Get-ADServerSettings).ViewEntireForest)
			{
				Set-ADServerSettings -ViewEntireForest $true -WarningAction SilentlyContinue
			}
			
			
			#...................................
			# Script
			#...................................
			
			#Check if a single server was specified
			if ($server)
			{
				#Run for single specified server
				[bool]$NoDAG = $true
				Write-Verbose $string20
				if ($Log) { Write-Logfile $string20 }
				try
				{
					$exchangeservers = Get-ExchangeServer $server -ErrorAction STOP
				}
				catch
				{
					#Exit because single server name was specified and couldn't be found in the organization
					Write-Verbose $string21
					if ($Log) { Write-Logfile $string21 }
					Write-Error $_.Exception.Message
					EXIT
				}
			}
			elseif ($serverlist)
			{
				#Run for a list of servers in a text file
				[bool]$NoDAG = $true
				Write-Verbose $string23
				if ($Log) { Write-Logfile $string23 }
				try
				{
					$tmpservers = @(Get-Content $serverlist -ErrorAction STOP)
					$exchangeservers = @($tmpservers | Get-ExchangeServer)
				}
				catch
				{
					#Exit because file could not be found
					Write-Verbose $string24
					if ($Log) { Write-Logfile $string24 }
					Write-Error $_.Exception.Message
					EXIT
				}
			}
			else
			{
				#This is the list of servers, DAGs, and databases to never alert for
				try
				{
					$ignorelist = @(Get-Content $ignorelistfile -ErrorAction STOP)
					if ($Log) { Write-Logfile $string28 }
					if ($Log)
					{
						if ($($ignorelist.count) -gt 0)
						{
							foreach ($line in $ignorelist)
							{
								Write-Logfile "- $line"
							}
						}
						else
						{
							Write-Logfile $string38
						}
					}
				}
				catch
				{
					Write-Warning $string22
					if ($Log) { Write-Logfile $string22 }
				}
				
				#Get all servers
				Write-Verbose $string25
				if ($Log) { Write-Logfile $string25 }
				$GetExchangeServerResults = @(Get-ExchangeServer | Sort-Object site, name)
				
				#Remove the servers that are ignored from the list of servers to check
				Write-Verbose $string26
				if ($Log) { Write-Logfile $string26 }
				foreach ($tmpserver in $GetExchangeServerResults)
				{
					if (!($ignorelist -icontains $tmpserver.name))
					{
						$exchangeservers = $exchangeservers += $tmpserver.identity
					}
				}
				
				if ($Log) { Write-Logfile $string29 }
				if ($Log)
				{
					foreach ($server in $exchangeservers)
					{
						Write-Logfile "- $server"
					}
				}
			}
			
			### Check if any Exchange 2013 servers exist
			if ($GetExchangeServerResults | Where-Object { $_.AdminDisplayVersion -like "Version 15.*" })
			{
				[bool]$HasE15 = $true
			}
			
			### Begin the Exchange Server health checks
			Write-Verbose $string27
			if ($Log) { Write-Logfile $string27 }
			foreach ($server in $exchangeservers)
			{
				Write-Host -ForegroundColor White "$string3 $server"
				if ($Log) { Write-Logfile "$string3 $server" }
				
				#Find out some details about the server
				try
				{
					$serverinfo = Get-ExchangeServer $server -ErrorAction Stop
				}
				catch
				{
					Write-Warning $_.Exception.Message
					if ($Log) { Write-Logfile $_.Exception.Message }
					$serverinfo = $null
				}
				
				if ($serverinfo -eq $null)
				{
					#Server is not an Exchange server
					Write-Host -ForegroundColor $warn $string0
					if ($Log) { Write-Logfile $string0 }
				}
				elseif ($serverinfo.IsEdgeServer)
				{
					Write-Host -ForegroundColor White $string8
					if ($Log) { Write-Logfile $string8 }
				}
				else
				{
					#Server is an Exchange server, continue the health check
					
					#Custom object properties
					$serverObj = New-Object PSObject
					$serverObj | Add-Member NoteProperty -Name "Server" -Value $server
					
					#Skip Site attribute for Exchange 2003 servers
					if ($serverinfo.AdminDisplayVersion -like "Version 6.*")
					{
						$serverObj | Add-Member NoteProperty -Name "Site" -Value "n/a"
					}
					else
					{
						$site = ($serverinfo.site.ToString()).Split("/")
						$serverObj | Add-Member NoteProperty -Name "Site" -Value $site[-1]
					}
					
					#Null and n/a the rest, will be populated as script progresses
					$serverObj | Add-Member NoteProperty -Name "DNS" -Value $null
					$serverObj | Add-Member NoteProperty -Name "Ping" -Value $null
					$serverObj | Add-Member NoteProperty -Name "Uptime (hrs)" -Value $null
					$serverObj | Add-Member NoteProperty -Name "Version" -Value $null
					$serverObj | Add-Member NoteProperty -Name "Roles" -Value $null
					$serverObj | Add-Member NoteProperty -Name "Client Access Server Role Services" -Value "n/a"
					$serverObj | Add-Member NoteProperty -Name "Hub Transport Server Role Services" -Value "n/a"
					$serverObj | Add-Member NoteProperty -Name "Mailbox Server Role Services" -Value "n/a"
					$serverObj | Add-Member NoteProperty -Name "Unified Messaging Server Role Services" -Value "n/a"
					$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "n/a"
					$serverObj | Add-Member NoteProperty -Name "Queue Length" -Value "n/a"
					$serverObj | Add-Member NoteProperty -Name "PF DBs Mounted" -Value "n/a"
					$serverObj | Add-Member NoteProperty -Name "MB DBs Mounted" -Value "n/a"
					$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "n/a"
					$serverObj | Add-Member NoteProperty -Name "MAPI Test" -Value "n/a"
					
					#Check server name resolves in DNS
					if ($Log) { Write-Logfile $string30 }
					Write-Host "DNS Check: " -NoNewline;
					try
					{
						$ip = @([System.Net.Dns]::GetHostByName($server).AddressList | Select-Object IPAddressToString -ExpandProperty IPAddressToString)
					}
					catch
					{
						Write-Host -ForegroundColor $warn $_.Exception.Message
						if ($Log) { Write-Logfile $_.Exception.Message }
						$ip = $null
					}
					
					if ($ip -ne $null)
					{
						Write-Host -ForegroundColor $pass "Pass"
						if ($Log) { Write-Logfile $string31 }
						$serverObj | Add-Member NoteProperty -Name "DNS" -Value "Pass" -Force
						
						#Is server online
						if ($Log) { Write-Logfile $string32 }
						Write-Host "Ping Check: " -NoNewline;
						
						$ping = $null
						try
						{
							$ping = Test-Connection $server -Quiet -ErrorAction Stop
						}
						catch
						{
							Write-Host -ForegroundColor $warn $_.Exception.Message
							if ($Log) { Write-Logfile $_.Exception.Message }
						}
						
						switch ($ping)
						{
							$true {
								Write-Host -ForegroundColor $pass "Pass"
								$serverObj | Add-Member NoteProperty -Name "Ping" -Value "Pass" -Force
								if ($Log) { Write-Logfile $string33 }
							}
							default
							{
								Write-Host -ForegroundColor $fail "Fail"
								$serverObj | Add-Member NoteProperty -Name "Ping" -Value "Fail" -Force
								$serversummary += "$server - $string18"
								if ($Log) { Write-Logfile $string18 }
							}
						}
						
						#Uptime check, even if ping fails
						if ($Log) { Write-Logfile $string34 }
						[int]$uptime = $null
						#$laststart = $null
						$OS = $null
						
						try
						{
							#$laststart = [System.Management.ManagementDateTimeconverter]::ToDateTime((Get-WmiObject -Class Win32_OperatingSystem -computername $server -ErrorAction Stop).LastBootUpTime)
							$OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server -ErrorAction STOP
						}
						catch
						{
							Write-Host -ForegroundColor $warn $_.Exception.Message
							if ($Log) { Write-Logfile $_.Exception.Message }
						}
						
						Write-Host "Uptime (hrs): " -NoNewline
						
						if ($OS -eq $null)
						{
							[string]$uptime = $string17
							if ($Log) { Write-Logfile $string17 }
							switch ($ping)
							{
								$true { $serversummary += "$server - $string17" }
								default { $serversummary += "$server - $string17" }
							}
						}
						else
						{
							$timespan = $OS.ConvertToDateTime($OS.LocalDateTime) – $OS.ConvertToDateTime($OS.LastBootUpTime)
							[int]$uptime = "{0:00}" -f $timespan.TotalHours
							Switch ($uptime -gt 23)
							{
								$true { Write-Host -ForegroundColor $pass $uptime }
								$false { Write-Host -ForegroundColor $warn $uptime; $serversummary += "$server - Uptime is less than 24 hours" }
								default { Write-Host -ForegroundColor $warn $uptime; $serversummary += "$server - Uptime is less than 24 hours" }
							}
						}
						
						if ($Log) { Write-Logfile "Uptime is $uptime hours" }
						
						$serverObj | Add-Member NoteProperty -Name "Uptime (hrs)" -Value $uptime -Force
						
						if ($ping -or ($uptime -ne $string17))
						{
							#Determine the friendly version number
							$ExVer = $serverinfo.AdminDisplayVersion
							Write-Host "Server version: " -NoNewline;
							
							if ($ExVer -like "Version 6.*")
							{
								$version = "Exchange 2003"
							}
							
							if ($ExVer -like "Version 8.*")
							{
								$version = "Exchange 2007"
							}
							
							if ($ExVer -like "Version 14.*")
							{
								$version = "Exchange 2010"
							}
							
							if ($ExVer -like "Version 15.0*")
							{
								$version = "Exchange 2013"
							}
							
							if ($ExVer -like "Version 15.1*")
							{
								$version = "Exchange 2016"
							}
							
							Write-Host $version
							if ($Log) { Write-Logfile "Server is running $version" }
							$serverObj | Add-Member NoteProperty -Name "Version" -Value $version -Force
							
							if ($version -eq "Exchange 2003")
							{
								Write-Host $string12
								if ($Log) { Write-Logfile $string12 }
							}
							
							#START - Exchange 2013/2010/2007 Health Checks
							if ($version -ne "Exchange 2003")
							{
								Write-Host "Roles:" $serverinfo.ServerRole
								if ($Log) { Write-Logfile "Server roles: $($serverinfo.ServerRole)" }
								$serverObj | Add-Member NoteProperty -Name "Roles" -Value $serverinfo.ServerRole -Force
								
								$IsEdge = $serverinfo.IsEdgeServer
								$IsHub = $serverinfo.IsHubTransportServer
								$IsCAS = $serverinfo.IsClientAccessServer
								$IsMB = $serverinfo.IsMailboxServer
								
								#START - General Server Health Check
								#Skipping Edge Transports for the general health check, as firewalls usually get
								#in the way. If you want to include them, remove this If.
								if ($IsEdge -ne $true)
								{
									#Service health is an array due to how multi-role servers return Test-ServiceHealth status
									if ($Log) { Write-Logfile $string35 }
									$servicehealth = @()
									$e15casservicehealth = @()
									try
									{
										$servicehealth = @(Test-ServiceHealth $server -ErrorAction Stop)
									}
									catch
									{
										#Workaround for Test-ServiceHealth problem with CAS-only Exchange 2013 servers
										#More info: http://exchangeserverpro.com/exchange-2013-test-servicehealth-error/
										if ($_.Exception.Message -like "*There are no Microsoft Exchange 2007 server roles installed*")
										{
											if ($Log) { Write-Logfile $string52 }
											$e15casservicehealth = Test-E15CASServiceHealth($server)
										}
										else
										{
											$serversummary += "$server - $string4"
											Write-Host -ForegroundColor $warn $string4 ":" $_.Exception
											if ($Log) { Write-Logfile $_.Exception }
											$serverObj | Add-Member NoteProperty -Name "Client Access Server Role Services" -Value $string4 -Force
											$serverObj | Add-Member NoteProperty -Name "Hub Transport Server Role Services" -Value $string4 -Force
											$serverObj | Add-Member NoteProperty -Name "Mailbox Server Role Services" -Value $string4 -Force
											$serverObj | Add-Member NoteProperty -Name "Unified Messaging Server Role Services" -Value $string4 -Force
										}
									}
									
									if ($servicehealth)
									{
										foreach ($s in $servicehealth)
										{
											$roleName = $s.Role
											Write-Host $roleName "Services: " -NoNewline;
											
											switch ($s.RequiredServicesRunning)
											{
												$true {
													$svchealth = "Pass"
													Write-Host -ForegroundColor $pass "Pass"
												}
												$false {
													$svchealth = "Fail"
													Write-Host -ForegroundColor $fail "Fail"
													$serversummary += "$server - $rolename $string5"
												}
												default
												{
													$svchealth = "Warn"
													Write-Host -ForegroundColor $warn "Warning"
													$serversummary += "$server - $rolename $string5"
												}
											}
											
											switch ($s.Role)
											{
												$casrole { $serverinfoservices = "Client Access Server Role Services" }
												$htrole { $serverinfoservices = "Hub Transport Server Role Services" }
												$mbrole { $serverinfoservices = "Mailbox Server Role Services" }
												$umrole { $serverinfoservices = "Unified Messaging Server Role Services" }
											}
											if ($Log) { Write-Logfile "$serverinfoservices status is $svchealth" }
											$serverObj | Add-Member NoteProperty -Name $serverinfoservices -Value $svchealth -Force
										}
									}
									
									if ($e15casservicehealth)
									{
										$serverinfoservices = "Client Access Server Role Services"
										if ($Log) { Write-Logfile "$serverinfoservices status is $e15casservicehealth" }
										$serverObj | Add-Member NoteProperty -Name $serverinfoservices -Value $e15casservicehealth -Force
										Write-Host $serverinfoservices ": " -NoNewline;
										switch ($e15casservicehealth)
										{
											"Pass" { Write-Host -ForegroundColor $pass "Pass" }
											"Fail" { Write-Host -ForegroundColor $fail "Fail" }
										}
									}
								}
								#END - General Server Health Check
								
								#START - Hub Transport Server Check
								if ($IsHub)
								{
									$q = $null
									if ($Log) { Write-Logfile $string36 }
									Write-Host "Total Queue: " -NoNewline;
									try
									{
										$q = Get-Queue -server $server -ErrorAction Stop | Where-Object { $_.DeliveryType -ne "ShadowRedundancy" }
									}
									catch
									{
										$serversummary += "$server - $string6"
										Write-Host -ForegroundColor $warn $string6
										Write-Warning $_.Exception.Message
										if ($Log) { Write-Logfile $string6 }
										if ($Log) { Write-Logfile $_.Exception.Message }
									}
									
									if ($q)
									{
										$qcount = $q | Measure-Object MessageCount -Sum
										[int]$qlength = $qcount.sum
										$serverObj | Add-Member NoteProperty -Name "Queue Length" -Value $qlength -Force
										if ($Log) { Write-Logfile "Queue length is $qlength" }
										if ($qlength -le $transportqueuewarn)
										{
											Write-Host -ForegroundColor $pass $qlength
											$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "Pass ($qlength)" -Force
										}
										elseif ($qlength -gt $transportqueuewarn -and $qlength -lt $transportqueuehigh)
										{
											Write-Host -ForegroundColor $warn $qlength
											$serversummary += "$server - Transport queue is above warning threshold"
											$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "Warn ($qlength)" -Force
										}
										else
										{
											Write-Host -ForegroundColor $fail $qlength
											$serversummary += "$server - Transport queue is above high threshold"
											$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "Fail ($qlength)" -Force
										}
									}
									else
									{
										$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "Unknown" -Force
									}
								}
								#END - Hub Transport Server Check
								
								#START - Mailbox Server Check
								if ($IsMB)
								{
									if ($Log) { Write-Logfile $string37 }
									
									#Get the PF and MB databases
									[array]$pfdbs = @(Get-PublicFolderDatabase -server $server -status -WarningAction SilentlyContinue)
									[array]$mbdbs = @(Get-MailboxDatabase -server $server -status | Where-Object { $_.Recovery -ne $true })
									
									if ($version -ne "Exchange 2007")
									{
										[array]$activedbs = @(Get-MailboxDatabase -server $server -status | Where-Object { $_.Recovery -ne $true -and $_.MountedOnServer -eq ($serverinfo.fqdn) })
									}
									else
									{
										[array]$activedbs = $mbdbs
									}
									
									#START - Database Mount Check
									
									#Check public folder databases
									if ($pfdbs.count -gt 0)
									{
										if ($Log) { Write-Logfile $string39 }
										Write-Host "Public Folder databases mounted: " -NoNewline;
										[string]$pfdbstatus = "Pass"
										[array]$alertdbs = @()
										foreach ($db in $pfdbs)
										{
											if (($db.mounted) -ne $true)
											{
												$pfdbstatus = "Fail"
												$alertdbs += $db.name
											}
										}
										
										$serverObj | Add-Member NoteProperty -Name "PF DBs Mounted" -Value $pfdbstatus -Force
										if ($Log) { Write-Logfile "$string40 $pfdbstatus" }
										
										if ($alertdbs.count -eq 0)
										{
											Write-Host -ForegroundColor $pass $pfdbstatus
										}
										else
										{
											Write-Host -ForegroundColor $fail $pfdbstatus
											$serversummary += "$server - $string7"
											Write-Host "Offline databases:"
											foreach ($al in $alertdbs)
											{
												Write-Host -ForegroundColor $fail `t$al
											}
										}
									}
									
									#Check mailbox databases
									if ($mbdbs.count -gt 0)
									{
										if ($Log) { Write-Logfile $string41 }
										
										[string]$mbdbstatus = "Pass"
										[array]$alertdbs = @()
										
										Write-Host "Mailbox databases mounted: " -NoNewline;
										foreach ($db in $mbdbs)
										{
											if (($db.mounted) -ne $true)
											{
												$mbdbstatus = "Fail"
												$alertdbs += $db.name
											}
										}
										
										$serverObj | Add-Member NoteProperty -Name "MB DBs Mounted" -Value $mbdbstatus -Force
										if ($Log) { Write-Logfile "$string42 $mbdbstatus" }
										
										if ($alertdbs.count -eq 0)
										{
											Write-Host -ForegroundColor $pass $mbdbstatus
										}
										else
										{
											$serversummary += "$server - $string9"
											Write-Host -ForegroundColor $fail $mbdbstatus
											Write-Host $string43
											if ($Log) { Write-Logfile $string43 }
											foreach ($al in $alertdbs)
											{
												Write-Host -ForegroundColor $fail `t$al
												if ($Log) { Write-Logfile "- $al" }
											}
										}
									}
									
									#END - Database Mount Check
									
									#START - MAPI Connectivity Test
									if ($activedbs.count -gt 0 -or $pfdbs.count -gt 0 -or $version -eq "Exchange 2007")
									{
										[string]$mapiresult = "Unknown"
										[array]$alertdbs = @()
										if ($Log) { Write-Logfile $string44 }
										Write-Host "MAPI connectivity: " -NoNewline;
										foreach ($db in $mbdbs)
										{
											$mapistatus = Test-MapiConnectivity -Database $db.Identity -PerConnectionTimeout $mapitimeout
											if ($mapistatus.Result.Value -eq $null)
											{
												$mapiresult = $mapistatus.Result
											}
											else
											{
												$mapiresult = $mapistatus.Result.Value
											}
											if (($mapiresult) -ne "Success")
											{
												$mapistatus = "Fail"
												$alertdbs += $db.name
											}
										}
										
										$serverObj | Add-Member NoteProperty -Name "MAPI Test" -Value $mapiresult -Force
										if ($Log) { Write-Logfile "$string45  $mapiresult" }
										
										if ($alertdbs.count -eq 0)
										{
											Write-Host -ForegroundColor $pass  $mapiresult
										}
										else
										{
											$serversummary += "$server - $string10"
											Write-Host -ForegroundColor $fail  $mapiresult
											Write-Host $string46
											if ($Log) { Write-Logfile $string46 }
											foreach ($al in $alertdbs)
											{
												Write-Host -ForegroundColor $fail `t$al
												if ($Log) { Write-Logfile "- $al" }
											}
										}
									}
									#END - MAPI Connectivity Test
									
									#START - Mail Flow Test
									if ($version -eq "Exchange 2007" -and $mbdbs.count -gt 0 -and $HasE15)
									{
										#Skip Exchange 2007 mail flow tests when run from Exchange 2013
										if ($Log) { Write-Logfile $string47 }
										Write-Host "Mail flow test: Skipped"
										$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value $string51 -Force
										if ($Log) { Write-Logfile $string51 }
									}
									elseif ($activedbs.count -gt 0 -and $HasE15)
									{
										if ($Log) { Write-Logfile $string47 }
										Write-Host "Mail flow test: " -NoNewline;
										$e15mailflowresult = Test-E15MailFlow($Server)
										$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value $e15mailflowresult -Force
										if ($Log) { Write-Logfile "$string48 $e15mailflowresult" }
										
										if ($e15mailflowresult -eq $success)
										{
											Write-Host -ForegroundColor $pass $e15mailflowresult
											$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "Pass" -Force
										}
										else
										{
											$serversummary += "$server - $string11"
											Write-Host -ForegroundColor $fail $e15mailflowresult
											$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "Fail" -Force
										}
									}
									elseif ($activedbs.count -gt 0 -or ($version -eq "Exchange 2007" -and $mbdbs.count -gt 0))
									{
										$flow = $null
										$testmailflowresult = $null
										
										if ($Log) { Write-Logfile $string47 }
										Write-Host "Mail flow test: " -NoNewline;
										try
										{
											$flow = Test-Mailflow $server -ErrorAction Stop
										}
										catch
										{
											$testmailflowresult = $_.Exception.Message
											if ($Log) { Write-Logfile $_.Exception.Message }
										}
										
										if ($flow)
										{
											$testmailflowresult = $flow.testmailflowresult
											if ($Log) { Write-Logfile "$string48 $testmailflowresult" }
										}
										
										if ($testmailflowresult -eq "Success" -or $testmailflowresult -eq $success)
										{
											Write-Host -ForegroundColor $pass $testmailflowresult
											$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "Pass" -Force
										}
										else
										{
											$serversummary += "$server - $string11"
											Write-Host -ForegroundColor $fail $testmailflowresult
											$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "Fail" -Force
										}
									}
									else
									{
										Write-Host "Mail flow test: No active mailbox databases"
										$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value $string49 -Force
										if ($Log) { Write-Logfile $string49 }
									}
									#END - Mail Flow Test
								}
								#END - Mailbox Server Check
								
							}
							#END - Exchange 2013/2010/2007 Health Checks
							if ($Log) { Write-Logfile "$string50 $server" }
							$report = $report + $serverObj
						}
						else
						{
							#Server is not reachable and uptime could not be retrieved
							Write-Host -ForegroundColor $warn $string1
							if ($Log) { Write-Logfile $string1 }
							$serversummary += "$server - $string1"
							$serverObj | Add-Member NoteProperty -Name "Ping" -Value "Fail" -Force
							if ($Log) { Write-Logfile "$string50 $server" }
							$report = $report + $serverObj
						}
					}
					else
					{
						Write-Host -ForegroundColor $Fail "Fail"
						Write-Host -ForegroundColor $warn $string13
						if ($Log) { Write-Logfile $string13 }
						$serversummary += "$server - $string13"
						$serverObj | Add-Member NoteProperty -Name "DNS" -Value "Fail" -Force
						if ($Log) { Write-Logfile "$string50 $server" }
						$report = $report + $serverObj
					}
				}
			}
			### End the Exchange Server health checks
			
			
			### Begin DAG Health Report
			
			#Check if -Server or -Serverlist parameter was used, and skip if it was
			if (!($NoDAG))
			{
				if ($Log) { Write-Logfile $string60 }
				Write-Verbose "Retrieving Database Availability Groups"
				
				#Get all DAGs
				$tmpdags = @(Get-DatabaseAvailabilityGroup)
				$tmpstring = "$($tmpdags.count) DAGs found"
				Write-Verbose $tmpstring
				if ($Log) { Write-Logfile $tmpstring }
				
				#Remove DAGs in ignorelist
				foreach ($tmpdag in $tmpdags)
				{
					if (!($ignorelist -icontains $tmpdag.name))
					{
						$dags += $tmpdag
					}
				}
				
				$tmpstring = "$($dags.count) DAGs will be checked"
				Write-Verbose $tmpstring
				if ($Log) { Write-Logfile $tmpstring }
				
				if ($Log) { Write-Logfile $string68 }
				if ($Log)
				{
					foreach ($dag in $dags)
					{
						Write-Logfile "- $dag"
					}
				}
			}
			
			if ($($dags.count) -gt 0)
			{
				foreach ($dag in $dags)
				{
					
					#Strings for use in the HTML report/email
					$dagsummaryintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Health Summary:</p>"
					$dagdetailintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Health Details:</p>"
					$dagmemberintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Member Health:</p>"
					
					$dagdbcopyReport = @() #Database copy health report
					$dagciReport = @() #Content Index health report
					$dagmemberReport = @() #DAG member server health report
					$dagdatabaseSummary = @() #Database health summary report
					$dagdatabases = @() #Array of databases in the DAG
					
					$tmpstring = "---- Processing DAG $($dag.Name)"
					Write-Verbose $tmpstring
					if ($Log) { Write-Logfile $tmpstring }
					
					$dagmembers = @($dag | Select-Object -ExpandProperty Servers | Sort-Object Name)
					$tmpstring = "$($dagmembers.count) DAG members found"
					Write-Verbose $tmpstring
					if ($Log) { Write-Logfile $tmpstring }
					
					#Get all databases in the DAG
					if ($HasE15)
					{
						$tmpdatabases = @(Get-MailboxDatabase -Status -IncludePreExchange2013 | Where-Object { $_.Recovery -ne $true -and $_.MasterServerOrAvailabilityGroup -eq $dag.Name } | Sort-Object Name)
					}
					else
					{
						$tmpdatabases = @(Get-MailboxDatabase -Status | Where-Object { $_.Recovery -ne $true -and $_.MasterServerOrAvailabilityGroup -eq $dag.Name } | Sort-Object Name)
					}
					
					foreach ($tmpdatabase in $tmpdatabases)
					{
						if (!($ignorelist -icontains $tmpdatabase.name))
						{
							$dagdatabases += $tmpdatabase
						}
					}
					
					$tmpstring = "$($dagdatabases.count) DAG databases will be checked"
					Write-Verbose $tmpstring
					if ($Log) { Write-Logfile $tmpstring }
					
					if ($Log) { Write-Logfile $string69 }
					if ($Log)
					{
						foreach ($database in $dagdatabases)
						{
							Write-Logfile "- $database"
						}
					}
					
					foreach ($database in $dagdatabases)
					{
						$tmpstring = "---- Processing database $database"
						Write-Verbose $tmpstring
						if ($Log) { Write-Logfile $tmpstring }
						
						$activationPref = $null
						$totalcopies = $null
						$healthycopies = $null
						$unhealthycopies = $null
						$healthyqueues = $null
						$unhealthyqueues = $null
						$laggedqueues = $null
						$healthyindexes = $null
						$unhealthyindexes = $null
						
						#Custom object for Database
						$objectHash = @{
							"Database" = $database.Identity
							"Mounted on" = "Unknown"
							"Preference" = $null
							"Total Copies" = $null
							"Healthy Copies" = $null
							"Unhealthy Copies" = $null
							"Healthy Queues" = $null
							"Unhealthy Queues" = $null
							"Lagged Queues" = $null
							"Healthy Indexes" = $null
							"Unhealthy Indexes" = $null
						}
						$databaseObj = New-Object PSObject -Property $objectHash
						
						$dbcopystatus = @($database | Get-MailboxDatabaseCopyStatus)
						$tmpstring = "$database has $($dbcopystatus.Count) copies"
						Write-Verbose $tmpstring
						if ($Log) { Write-Logfile $tmpstring }
						
						foreach ($dbcopy in $dbcopystatus)
						{
							#Custom object for DB copy
							$objectHash = @{
								"Database Copy" = $dbcopy.Identity
								"Database Name" = $dbcopy.DatabaseName
								"Mailbox Server" = $null
								"Activation Preference" = $null
								"Status"	    = $null
								"Copy Queue"    = $null
								"Replay Queue"  = $null
								"Replay Lagged" = $null
								"Truncation Lagged" = $null
								"Content Index" = $null
							}
							$dbcopyObj = New-Object PSObject -Property $objectHash
							
							$tmpstring = "Database Copy: $($dbcopy.Identity)"
							Write-Verbose $tmpstring
							if ($Log) { Write-Logfile $tmpstring }
							
							$mailboxserver = $dbcopy.MailboxServer
							$tmpstring = "Server: $mailboxserver"
							Write-Verbose $tmpstring
							if ($Log) { Write-Logfile $tmpstring }
							
							$pref = ($database | Select-Object -ExpandProperty ActivationPreference | Where-Object { $_.Key -ieq $mailboxserver }).Value
							$tmpstring = "Activation Preference: $pref"
							Write-Verbose $tmpstring
							if ($Log) { Write-Logfile $tmpstring }
							
							$copystatus = $dbcopy.Status
							$tmpstring = "Status: $copystatus"
							Write-Verbose $tmpstring
							if ($Log) { Write-Logfile $tmpstring }
							
							[int]$copyqueuelength = $dbcopy.CopyQueueLength
							$tmpstring = "Copy Queue: $copyqueuelength"
							Write-Verbose $tmpstring
							if ($Log) { Write-Logfile $tmpstring }
							
							[int]$replayqueuelength = $dbcopy.ReplayQueueLength
							$tmpstring = "Replay Queue: $replayqueuelength"
							Write-Verbose $tmpstring
							if ($Log) { Write-Logfile $tmpstring }
							
							if ($($dbcopy.ContentIndexErrorMessage -match "is disabled in Active Directory"))
							{
								$contentindexstate = "Disabled"
							}
							else
							{
								$contentindexstate = $dbcopy.ContentIndexState
							}
							$tmpstring = "Content Index: $contentindexstate"
							Write-Verbose $tmpstring
							if ($Log) { Write-Logfile $tmpstring }
							
							#Checking whether this is a replay lagged copy
							$replaylagcopies = @($database | Select-Object -ExpandProperty ReplayLagTimes | Where-Object { $_.Value -gt 0 })
							if ($($replaylagcopies.count) -gt 0)
							{
								[bool]$replaylag = $false
								foreach ($replaylagcopy in $replaylagcopies)
								{
									if ($replaylagcopy.Key -ieq $mailboxserver)
									{
										$tmpstring = "$database is replay lagged on $mailboxserver"
										Write-Verbose $tmpstring
										if ($Log) { Write-Logfile $tmpstring }
										[bool]$replaylag = $true
									}
								}
							}
							else
							{
								[bool]$replaylag = $false
							}
							$tmpstring = "Replay lag is $replaylag"
							Write-Verbose $tmpstring
							if ($Log) { Write-Logfile $tmpstring }
							
							#Checking for truncation lagged copies
							$truncationlagcopies = @($database | Select-Object -ExpandProperty TruncationLagTimes | Where-Object { $_.Value -gt 0 })
							if ($($truncationlagcopies.count) -gt 0)
							{
								[bool]$truncatelag = $false
								foreach ($truncationlagcopy in $truncationlagcopies)
								{
									if ($truncationlagcopy.Key -eq $mailboxserver)
									{
										$tmpstring = "$database is truncate lagged on $mailboxserver"
										Write-Verbose $tmpstring
										if ($Log) { Write-Logfile $tmpstring }
										[bool]$truncatelag = $true
									}
								}
							}
							else
							{
								[bool]$truncatelag = $false
							}
							$tmpstring = "Truncation lag is $truncatelag"
							Write-Verbose $tmpstring
							if ($Log) { Write-Logfile $tmpstring }
							
							$dbcopyObj | Add-Member NoteProperty -Name "Mailbox Server" -Value $mailboxserver -Force
							$dbcopyObj | Add-Member NoteProperty -Name "Activation Preference" -Value $pref -Force
							$dbcopyObj | Add-Member NoteProperty -Name "Status" -Value $copystatus -Force
							$dbcopyObj | Add-Member NoteProperty -Name "Copy Queue" -Value $copyqueuelength -Force
							$dbcopyObj | Add-Member NoteProperty -Name "Replay Queue" -Value $replayqueuelength -Force
							$dbcopyObj | Add-Member NoteProperty -Name "Replay Lagged" -Value $replaylag -Force
							$dbcopyObj | Add-Member NoteProperty -Name "Truncation Lagged" -Value $truncatelag -Force
							$dbcopyObj | Add-Member NoteProperty -Name "Content Index" -Value $contentindexstate -Force
							
							$dagdbcopyReport += $dbcopyObj
						}
						
						$copies = @($dagdbcopyReport | Where-Object { ($_."Database Name" -eq $database) })
						
						$mountedOn = ($copies | Where-Object { ($_.Status -eq "Mounted") })."Mailbox Server"
						if ($mountedOn)
						{
							$databaseObj | Add-Member NoteProperty -Name "Mounted on" -Value $mountedOn -Force
						}
						
						$activationPref = ($copies | Where-Object { ($_.Status -eq "Mounted") })."Activation Preference"
						$databaseObj | Add-Member NoteProperty -Name "Preference" -Value $activationPref -Force
						
						$totalcopies = $copies.count
						$databaseObj | Add-Member NoteProperty -Name "Total Copies" -Value $totalcopies -Force
						
						$healthycopies = @($copies | Where-Object { (($_.Status -eq "Mounted") -or ($_.Status -eq "Healthy")) }).Count
						$databaseObj | Add-Member NoteProperty -Name "Healthy Copies" -Value $healthycopies -Force
						
						$unhealthycopies = @($copies | Where-Object { (($_.Status -ne "Mounted") -and ($_.Status -ne "Healthy")) }).Count
						$databaseObj | Add-Member NoteProperty -Name "Unhealthy Copies" -Value $unhealthycopies -Force
						
						$healthyqueues = @($copies | Where-Object { (($_."Copy Queue" -lt $replqueuewarning) -and (($_."Replay Queue" -lt $replqueuewarning)) -and ($_."Replay Lagged" -eq $false)) }).Count
						$databaseObj | Add-Member NoteProperty -Name "Healthy Queues" -Value $healthyqueues -Force
						
						$unhealthyqueues = @($copies | Where-Object { (($_."Copy Queue" -ge $replqueuewarning) -or (($_."Replay Queue" -ge $replqueuewarning) -and ($_."Replay Lagged" -eq $false))) }).Count
						$databaseObj | Add-Member NoteProperty -Name "Unhealthy Queues" -Value $unhealthyqueues -Force
						
						$laggedqueues = @($copies | Where-Object { ($_."Replay Lagged" -eq $true) -or ($_."Truncation Lagged" -eq $true) }).Count
						$databaseObj | Add-Member NoteProperty -Name "Lagged Queues" -Value $laggedqueues -Force
						
						$healthyindexes = @($copies | Where-Object { ($_."Content Index" -eq "Healthy" -or $_."Content Index" -eq "Disabled" -or $_."Content Index" -eq "AutoSuspended") }).Count
						$databaseObj | Add-Member NoteProperty -Name "Healthy Indexes" -Value $healthyindexes -Force
						
						$unhealthyindexes = @($copies | Where-Object { ($_."Content Index" -ne "Healthy" -and $_."Content Index" -ne "Disabled" -and $_."Content Index" -ne "AutoSuspended") }).Count
						$databaseObj | Add-Member NoteProperty -Name "Unhealthy Indexes" -Value $unhealthyindexes -Force
						
						$dagdatabaseSummary += $databaseObj
						
					}
					
					#Get Test-Replication Health results for each DAG member
					foreach ($dagmember in $dagmembers)
					{
						$replicationhealth = $null
						
						$replicationhealthitems = @{
							ClusterService	     = $null
							ReplayService	     = $null
							ActiveManager	     = $null
							TasksRpcListener	 = $null
							TcpListener		     = $null
							ServerLocatorService = $null
							DagMembersUp		 = $null
							ClusterNetwork	     = $null
							QuorumGroup		     = $null
							FileShareQuorum	     = $null
							DatabaseRedundancy   = $null
							DatabaseAvailability = $null
							DBCopySuspended	     = $null
							DBCopyFailed		 = $null
							DBInitializing	     = $null
							DBDisconnected	     = $null
							DBLogCopyKeepingUp   = $null
							DBLogReplayKeepingUp = $null
						}
						
						$memberObj = New-Object PSObject -Property $replicationhealthitems
						$memberObj | Add-Member NoteProperty -Name "Server" -Value $($dagmember.Name)
						
						$tmpstring = "---- Checking replication health for $($dagmember.Name)"
						Write-Verbose $tmpstring
						if ($Log) { Write-Logfile $tmpstring }
						
						if ($HasE15)
						{
							$DagMemberVer = ($GetExchangeServerResults | Where-Object { $_.Name -ieq $dagmember.Name }).AdminDisplayVersion.ToString()
						}
						
						
						if ($DagMemberVer -like "Version 14.*")
						{
							if ($Log) { Write-Logfile "Using E14 replication health test workaround" }
							$replicationhealth = Test-E14ReplicationHealth $dagmember
						}
						else
						{
							$replicationhealth = Test-ReplicationHealth -Identity $dagmember
						}
						
						foreach ($healthitem in $replicationhealth)
						{
							if ($($healthitem.Result) -eq $null)
							{
								$healthitemresult = "n/a"
							}
							else
							{
								$healthitemresult = $($healthitem.Result)
							}
							$tmpstring = "$($healthitem.Check) $healthitemresult"
							Write-Verbose $tmpstring
							if ($Log) { Write-Logfile $tmpstring }
							$memberObj | Add-Member NoteProperty -Name $($healthitem.Check) -Value $healthitemresult -Force
						}
						$dagmemberReport += $memberObj
					}
					
					
					#Generate the HTML from the DAG health checks
					if ($SendEmail -or $ReportFile)
					{
						
						####Begin Summary Table HTML
						$dagdatabaseSummaryHtml = $null
						#Begin Summary table HTML header
						$htmltableheader = "<p>
                            <table>
                            <tr>
                            <th>Database</th>
                            <th>Mounted on</th>
                            <th>Preference</th>
                            <th>Total Copies</th>
                            <th>Healthy Copies</th>
                            <th>Unhealthy Copies</th>
                            <th>Healthy Queues</th>
                            <th>Unhealthy Queues</th>
                            <th>Lagged Queues</th>
                            <th>Healthy Indexes</th>
                            <th>Unhealthy Indexes</th>
                            </tr>"
						
						$dagdatabaseSummaryHtml += $htmltableheader
						#End Summary table HTML header
						
						#Begin Summary table HTML rows
						foreach ($line in $dagdatabaseSummary)
						{
							$htmltablerow = "<tr>"
							$htmltablerow += "<td><strong>$($line.Database)</strong></td>"
							
							#Warn if mounted server is still unknown
							switch ($($line."Mounted on"))
							{
								"Unknown" {
									$htmltablerow += "<td class=""warn"">$($line."Mounted on")</td>"
									$dagsummary += "$($line.Database) - $string61"
								}
								default { $htmltablerow += "<td>$($line."Mounted on")</td>" }
							}
							
							#Warn if DB is mounted on a server that is not Activation Preference 1
							if ($($line.Preference) -gt 1)
							{
								$htmltablerow += "<td class=""warn"">$($line.Preference)</td>"
								$dagsummary += "$($line.Database) - $string62 $($line.Preference)"
							}
							else
							{
								$htmltablerow += "<td class=""pass"">$($line.Preference)</td>"
							}
							
							$htmltablerow += "<td>$($line."Total Copies")</td>"
							
							#Show as info if health copies is 1 but total copies also 1,
							#Warn if healthy copies is 1, Fail if 0
							switch ($($line."Healthy Copies"))
							{
								0 { $htmltablerow += "<td class=""fail"">$($line."Healthy Copies")</td>" }
								1 {
									if ($($line."Total Copies") -eq $($line."Healthy Copies"))
									{
										$htmltablerow += "<td class=""info"">$($line."Healthy Copies")</td>"
									}
									else
									{
										$htmltablerow += "<td class=""warn"">$($line."Healthy Copies")</td>"
									}
								}
								default { $htmltablerow += "<td class=""pass"">$($line."Healthy Copies")</td>" }
							}
							
							#Warn if unhealthy copies is 1, fail if more than 1
							switch ($($line."Unhealthy Copies"))
							{
								0 { $htmltablerow += "<td class=""pass"">$($line."Unhealthy Copies")</td>" }
								1 {
									$htmltablerow += "<td class=""warn"">$($line."Unhealthy Copies")</td>"
									$dagsummary += "$($line.Database) - $string63 $($line."Unhealthy Copies") $string65 $($line."Total Copies") $string66"
								}
								default
								{
									$htmltablerow += "<td class=""fail"">$($line."Unhealthy Copies")</td>"
									$dagsummary += "$($line.Database) - $string63 $($line."Unhealthy Copies") $string65 $($line."Total Copies") $string66"
								}
							}
							
							#Warn if healthy queues + lagged queues is less than total copies
							#Fail if no healthy queues
							if ($($line."Total Copies") -eq ($($line."Healthy Queues") + $($line."Lagged Queues")))
							{
								$htmltablerow += "<td class=""pass"">$($line."Healthy Queues")</td>"
							}
							else
							{
								$dagsummary += "$($line.Database) - $string64 $($line."Healthy Queues") $string65 $($line."Total Copies") $string66"
								switch ($($line."Healthy Queues"))
								{
									0 { $htmltablerow += "<td class=""fail"">$($line."Healthy Queues")</td>" }
									default { $htmltablerow += "<td class=""warn"">$($line."Healthy Queues")</td>" }
								}
							}
							
							#Fail if unhealthy queues = total queues
							#Warn if more than one unhealthy queue
							if ($($line."Total Queues") -eq $($line."Unhealthy Queues"))
							{
								$htmltablerow += "<td class=""fail"">$($line."Unhealthy Queues")</td>"
							}
							else
							{
								switch ($($line."Unhealthy Queues"))
								{
									0 { $htmltablerow += "<td class=""pass"">$($line."Unhealthy Queues")</td>" }
									default { $htmltablerow += "<td class=""warn"">$($line."Unhealthy Queues")</td>" }
								}
							}
							
							#Info for lagged queues
							switch ($($line."Lagged Queues"))
							{
								0 { $htmltablerow += "<td>$($line."Lagged Queues")</td>" }
								default { $htmltablerow += "<td class=""info"">$($line."Lagged Queues")</td>" }
							}
							
							#Pass if healthy indexes = total copies
							#Warn if healthy indexes less than total copies
							#Fail if healthy indexes = 0
							if ($($line."Total Copies") -eq $($line."Healthy Indexes"))
							{
								$htmltablerow += "<td class=""pass"">$($line."Healthy Indexes")</td>"
							}
							else
							{
								$dagsummary += "$($line.Database) - $string67 $($line."Unhealthy Indexes") $string65 $($line."Total Copies") $string66"
								switch ($($line."Healthy Indexes"))
								{
									0 { $htmltablerow += "<td class=""fail"">$($line."Healthy Indexes")</td>" }
									default { $htmltablerow += "<td class=""warn"">$($line."Healthy Indexes")</td>" }
								}
							}
							
							#Fail if unhealthy indexes = total copies
							#Warn if unhealthy indexes 1 or more
							#Pass if unhealthy indexes = 0
							if ($($line."Total Copies") -eq $($line."Unhealthy Indexes"))
							{
								$htmltablerow += "<td class=""fail"">$($line."Unhealthy Indexes")</td>"
							}
							else
							{
								switch ($($line."Unhealthy Indexes"))
								{
									0 { $htmltablerow += "<td class=""pass"">$($line."Unhealthy Indexes")</td>" }
									default { $htmltablerow += "<td class=""warn"">$($line."Unhealthy Indexes")</td>" }
								}
							}
							
							$htmltablerow += "</tr>"
							$dagdatabaseSummaryHtml += $htmltablerow
						}
						$dagdatabaseSummaryHtml += "</table>
                                    </p>"
						#End Summary table HTML rows
						####End Summary Table HTML
						
						####Begin Detail Table HTML
						$databasedetailsHtml = $null
						#Begin Detail table HTML header
						$htmltableheader = "<p>
                                <table>
                                <tr>
                                <th>Database Copy</th>
                                <th>Database Name</th>
                                <th>Mailbox Server</th>
                                <th>Activation Preference</th>
                                <th>Status</th>
                                <th>Copy Queue</th>
                                <th>Replay Queue</th>
                                <th>Replay Lagged</th>
                                <th>Truncation Lagged</th>
                                <th>Content Index</th>
                                </tr>"
						
						$databasedetailsHtml += $htmltableheader
						#End Detail table HTML header
						
						#Begin Detail table HTML rows
						foreach ($line in $dagdbcopyReport)
						{
							$htmltablerow = "<tr>"
							$htmltablerow += "<td><strong>$($line."Database Copy")</strong></td>"
							$htmltablerow += "<td>$($line."Database Name")</td>"
							$htmltablerow += "<td>$($line."Mailbox Server")</td>"
							$htmltablerow += "<td>$($line."Activation Preference")</td>"
							
							Switch ($($line."Status"))
							{
								"Healthy" { $htmltablerow += "<td class=""pass"">$($line."Status")</td>" }
								"Mounted" { $htmltablerow += "<td class=""pass"">$($line."Status")</td>" }
								"Failed" { $htmltablerow += "<td class=""fail"">$($line."Status")</td>" }
								"FailedAndSuspended" { $htmltablerow += "<td class=""fail"">$($line."Status")</td>" }
								"ServiceDown" { $htmltablerow += "<td class=""fail"">$($line."Status")</td>" }
								"Dismounted" { $htmltablerow += "<td class=""fail"">$($line."Status")</td>" }
								default { $htmltablerow += "<td class=""warn"">$($line."Status")</td>" }
							}
							
							if ($($line."Copy Queue") -lt $replqueuewarning)
							{
								$htmltablerow += "<td class=""pass"">$($line."Copy Queue")</td>"
							}
							else
							{
								$htmltablerow += "<td class=""warn"">$($line."Copy Queue")</td>"
							}
							
							if (($($line."Replay Queue") -lt $replqueuewarning) -or ($($line."Replay Lagged") -eq $true))
							{
								$htmltablerow += "<td class=""pass"">$($line."Replay Queue")</td>"
							}
							else
							{
								$htmltablerow += "<td class=""warn"">$($line."Replay Queue")</td>"
							}
							
							
							Switch ($($line."Replay Lagged"))
							{
								$true { $htmltablerow += "<td class=""info"">$($line."Replay Lagged")</td>" }
								default { $htmltablerow += "<td>$($line."Replay Lagged")</td>" }
							}
							
							Switch ($($line."Truncation Lagged"))
							{
								$true { $htmltablerow += "<td class=""info"">$($line."Truncation Lagged")</td>" }
								default { $htmltablerow += "<td>$($line."Truncation Lagged")</td>" }
							}
							
							Switch ($($line."Content Index"))
							{
								"Healthy" { $htmltablerow += "<td class=""pass"">$($line."Content Index")</td>" }
								"Disabled" { $htmltablerow += "<td class=""info"">$($line."Content Index")</td>" }
								default { $htmltablerow += "<td class=""warn"">$($line."Content Index")</td>" }
							}
							
							$htmltablerow += "</tr>"
							$databasedetailsHtml += $htmltablerow
						}
						$databasedetailsHtml += "</table>
                                    </p>"
						#End Detail table HTML rows
						####End Detail Table HTML
						
						
						####Begin Member Table HTML
						$dagmemberHtml = $null
						#Begin Member table HTML header
						$htmltableheader = "<p>
                                <table>
                                <tr>
                                <th>Server</th>
                                <th>Cluster Service</th>
                                <th>Replay Service</th>
                                <th>Active Manager</th>
                                <th>Tasks RPC Listener</th>
                                <th>TCP Listener</th>
                                <th>Server Locator Service</th>
                                <th>DAG Members Up</th>
                                <th>Cluster Network</th>
                                <th>Quorum Group</th>
                                <th>File Share Quorum</th>
                                <th>Database Redundancy</th>
                                <th>Database Availability</th>
                                <th>DB Copy Suspended</th>
                                <th>DB Copy Failed</th>
                                <th>DB Initializing</th>
                                <th>DB Disconnected</th>
                                <th>DB Log Copy Keeping Up</th>
                                <th>DB Log Replay Keeping Up</th>
                                </tr>"
						
						$dagmemberHtml += $htmltableheader
						#End Member table HTML header
						
						#Begin Member table HTML rows
						foreach ($line in $dagmemberReport)
						{
							$htmltablerow = "<tr>"
							$htmltablerow += "<td><strong>$($line."Server")</strong></td>"
							$htmltablerow += (New-DAGMemberHTMLTableCell "ClusterService")
							$htmltablerow += (New-DAGMemberHTMLTableCell "ReplayService")
							$htmltablerow += (New-DAGMemberHTMLTableCell "ActiveManager")
							$htmltablerow += (New-DAGMemberHTMLTableCell "TasksRPCListener")
							$htmltablerow += (New-DAGMemberHTMLTableCell "TCPListener")
							$htmltablerow += (New-DAGMemberHTMLTableCell "ServerLocatorService")
							$htmltablerow += (New-DAGMemberHTMLTableCell "DAGMembersUp")
							$htmltablerow += (New-DAGMemberHTMLTableCell "ClusterNetwork")
							$htmltablerow += (New-DAGMemberHTMLTableCell "QuorumGroup")
							$htmltablerow += (New-DAGMemberHTMLTableCell "FileShareQuorum")
							$htmltablerow += (New-DAGMemberHTMLTableCell "DatabaseRedundancy")
							$htmltablerow += (New-DAGMemberHTMLTableCell "DatabaseAvailability")
							$htmltablerow += (New-DAGMemberHTMLTableCell "DBCopySuspended")
							$htmltablerow += (New-DAGMemberHTMLTableCell "DBCopyFailed")
							$htmltablerow += (New-DAGMemberHTMLTableCell "DBInitializing")
							$htmltablerow += (New-DAGMemberHTMLTableCell "DBDisconnected")
							$htmltablerow += (New-DAGMemberHTMLTableCell "DBLogCopyKeepingUp")
							$htmltablerow += (New-DAGMemberHTMLTableCell "DBLogReplayKeepingUp")
							$htmltablerow += "</tr>"
							$dagmemberHtml += $htmltablerow
						}
						$dagmemberHtml += "</table>
            </p>"
					}
					
					#Output the report objects to console, and optionally to email and HTML file
					#Forcing table format for console output due to issue with multiple output
					#objects that have different layouts
					
					#Write-Host "---- Database Copy Health Summary ----"
					#$dagdatabaseSummary | ft
					
					#Write-Host "---- Database Copy Health Details ----"
					#$dagdbcopyReport | ft
					
					#Write-Host "`r`n---- Server Test-Replication Report ----`r`n"
					#$dagmemberReport | ft
					
					if ($SendEmail -or $ReportFile)
					{
						$dagreporthtml = $dagsummaryintro + $dagdatabaseSummaryHtml + $dagdetailintro + $databasedetailsHtml + $dagmemberintro + $dagmemberHtml
						$dagreportbody += $dagreporthtml
					}
					
				}
			}
			else
			{
				$tmpstring = "No DAGs found"
				if ($Log) { Write-LogFile $tmpstring }
				Write-Verbose $tmpstring
				$dagreporthtml = "<p>No database availability groups found.</p>"
			}
			### End DAG Health Report
			
			Write-Host $string16
			### Begin report generation
			if ($ReportMode -or $SendEmail)
			{
				#Get report generation timestamp
				$reportime = Get-Date
				
				#Create HTML Report
				#Common HTML head and styles
				$htmlhead = "<html>
                <style>
                BODY{font-family: Arial; font-size: 8pt;}
                H1{font-size: 16px;}
                H2{font-size: 14px;}
                H3{font-size: 12px;}
                TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
                TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
                TD{border: 1px solid black; padding: 5px; }
                td.pass{background: #7FFF00;}
                td.warn{background: #FFE600;}
                td.fail{background: #FF0000; color: #ffffff;}
                td.info{background: #85D4FF;}
                </style>
                <body>
                <h1 align=""center"">Exchange Server Health Check Report</h1>
                <h3 align=""center"">Generated: $reportime</h3>"
				
				#Check if the server summary has 1 or more entries
				if ($($serversummary.count) -gt 0)
				{
					#Set alert flag to true
					$alerts = $true
					
					#Generate the HTML
					$serversummaryhtml = "<h3>Exchange Server Health Check Summary</h3>
                        <p>The following server errors and warnings were detected.</p>
                        <p>
                        <ul>"
					foreach ($reportline in $serversummary)
					{
						$serversummaryhtml += "<li>$reportline</li>"
					}
					$serversummaryhtml += "</ul></p>"
					$alerts = $true
				}
				else
				{
					#Generate the HTML to show no alerts
					$serversummaryhtml = "<h3>Exchange Server Health Check Summary</h3>
                        <p>No Exchange server health errors or warnings.</p>"
				}
				
				#Check if the DAG summary has 1 or more entries
				if ($($dagsummary.count) -gt 0)
				{
					#Set alert flag to true
					$alerts = $true
					
					#Generate the HTML
					$dagsummaryhtml = "<h3>Database Availability Group Health Check Summary</h3>
                        <p>The following DAG errors and warnings were detected.</p>
                        <p>
                        <ul>"
					foreach ($reportline in $dagsummary)
					{
						$dagsummaryhtml += "<li>$reportline</li>"
					}
					$dagsummaryhtml += "</ul></p>"
					$alerts = $true
				}
				else
				{
					#Generate the HTML to show no alerts
					$dagsummaryhtml = "<h3>Database Availability Group Health Check Summary</h3>
                        <p>No Exchange DAG errors or warnings.</p>"
				}
				
				
				#Exchange Server Health Report Table Header
				$htmltableheader = "<h3>Exchange Server Health</h3>
                        <p>
                        <table>
                        <tr>
                        <th>Server</th>
                        <th>Site</th>
                        <th>Roles</th>
                        <th>Version</th>
                        <th>DNS</th>
                        <th>Ping</th>
                        <th>Uptime (hrs)</th>
                        <th>Client Access Server Role Services</th>
                        <th>Hub Transport Server Role Services</th>
                        <th>Mailbox Server Role Services</th>
                        <th>Unified Messaging Server Role Services</th>
                        <th>Transport Queue</th>
                        <th>PF DBs Mounted</th>
                        <th>MB DBs Mounted</th>
                        <th>MAPI Test</th>
                        <th>Mail Flow Test</th>
                        </tr>"
				
				#Exchange Server Health Report Table
				$serverhealthhtmltable = $serverhealthhtmltable + $htmltableheader
				
				foreach ($reportline in $report)
				{
					$htmltablerow = "<tr>"
					$htmltablerow += "<td>$($reportline.server)</td>"
					$htmltablerow += "<td>$($reportline.site)</td>"
					$htmltablerow += "<td>$($reportline.roles)</td>"
					$htmltablerow += "<td>$($reportline.version)</td>"
					$htmltablerow += (New-ServerHealthHTMLTableCell "dns")
					$htmltablerow += (New-ServerHealthHTMLTableCell "ping")
					
					if ($($reportline."uptime (hrs)") -eq "Access Denied")
					{
						$htmltablerow += "<td class=""warn"">Access Denied</td>"
					}
					elseif ($($reportline."uptime (hrs)") -eq $string17)
					{
						$htmltablerow += "<td class=""warn"">$string17</td>"
					}
					else
					{
						$hours = [int]$($reportline."uptime (hrs)")
						if ($hours -le 24)
						{
							$htmltablerow += "<td class=""warn"">$hours</td>"
						}
						else
						{
							$htmltablerow += "<td class=""pass"">$hours</td>"
						}
					}
					
					$htmltablerow += (New-ServerHealthHTMLTableCell "Client Access Server Role Services")
					$htmltablerow += (New-ServerHealthHTMLTableCell "Hub Transport Server Role Services")
					$htmltablerow += (New-ServerHealthHTMLTableCell "Mailbox Server Role Services")
					$htmltablerow += (New-ServerHealthHTMLTableCell "Unified Messaging Server Role Services")
					#$htmltablerow += (New-ServerHealthHTMLTableCell "Transport Queue")
					if ($($reportline."Transport Queue") -match "Pass")
					{
						$htmltablerow += "<td class=""pass"">$($reportline."Transport Queue")</td>"
					}
					elseif ($($reportline."Transport Queue") -match "Warn")
					{
						$htmltablerow += "<td class=""warn"">$($reportline."Transport Queue")</td>"
					}
					elseif ($($reportline."Transport Queue") -match "Fail")
					{
						$htmltablerow += "<td class=""fail"">$($reportline."Transport Queue")</td>"
					}
					elseif ($($reportline."Transport Queue") -eq "n/a")
					{
						$htmltablerow += "<td>$($reportline."Transport Queue")</td>"
					}
					else
					{
						$htmltablerow += "<td class=""warn"">$($reportline."Transport Queue")</td>"
					}
					$htmltablerow += (New-ServerHealthHTMLTableCell "PF DBs Mounted")
					$htmltablerow += (New-ServerHealthHTMLTableCell "MB DBs Mounted")
					$htmltablerow += (New-ServerHealthHTMLTableCell "MAPI Test")
					$htmltablerow += (New-ServerHealthHTMLTableCell "Mail Flow Test")
					$htmltablerow += "</tr>"
					
					$serverhealthhtmltable = $serverhealthhtmltable + $htmltablerow
				}
				
				$serverhealthhtmltable = $serverhealthhtmltable + "</table></p>"
				
				$htmltail = "</body>
                </html>"
				
				$htmlreport = $htmlhead + $serversummaryhtml + $dagsummaryhtml + $serverhealthhtmltable + $dagreportbody + $htmltail
				
				if ($ReportMode -or $ReportFile)
				{
					$htmlreport | Out-File $ReportFile -Encoding UTF8
				}
				
				if ($SendEmail)
				{
					if ($alerts -eq $false -and $AlertsOnly -eq $true)
					{
						#Do not send email message
						Write-Host $string19
						if ($Log) { Write-Logfile $string19 }
					}
					else
					{
						#Send email message
						Write-Host $string14
						Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
					}
				}
			}
			### End report generation
			
			
			Write-Host $string15
			if ($Log) { Write-Logfile $string15 }
			
			
			$Reboot = $true
			
		}
		
		Function generateEASDeviceStats
		{
			#      Generate Reports for Exchange ActiveSync Device Statistics
                           <#
.SYNOPSIS
Get-EASDeviceReport.ps1 - Exchange Server ActiveSync device report

.DESCRIPTION 
Produces a report of ActiveSync device associations in the organization.

.OUTPUTS
Results are output to screen, as well as optional log file, HTML report, and HTML email

.PARAMETER SendEmail
Sends the HTML report via email using the SMTP configuration within the script.

.EXAMPLE
.\Get-EASDeviceReport.ps1
Produces a CSV file containing stats for all ActiveSync devices.

.EXAMPLE
.\Get-EASDeviceReport.ps1 -SendEmail -MailFrom:exchangeserver@exchangeserverpro.net -MailTo:paul@exchangeserverpro.com -MailServer:smtp.exchangeserverpro.net
Sends an email report with CSV file attached for all ActiveSync devices.

.EXAMPLE
.\Get-EASDeviceReport.ps1 -Age 30
Limits the report to devices that have not attempted synced in more than 30 days.

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:   http://paulcunningham.me
* Twitter:   https://twitter.com/paulcunningham
* LinkedIn:  http://au.linkedin.com/in/cunninghamp/
* Github:    https://github.com/cunninghamp

For more Exchange Server tips, tricks and news
check out Exchange Server Pro.

* Website:   http://exchangeserverpro.com
* Twitter:   http://twitter.com/exchservpro

License:

The MIT License (MIT)

Copyright (c) 2015 Paul Cunningham

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Change Log:
V1.00, 25/11/2013 - Initial version
V1.01, 11/02/2014 - Added parameters for emailing the report and specifying an "age" to report on
V1.02, 17/02/2014 - Fixed missing $mydir variable and added UTF8 encoding to Export-CSV and Send-MailMessage
V1.03, 19/02/2016 - Added OrganizationalUnit to report, plus minor fixes
#>
			
			
			#region user input
			Do
			{
				$smail = Read-Host "Do you want the report to be sent via e-mail? [y/n/c]"
			}
			until
			($smail -eq "y" -or $smail -eq "n" -or $smail -eq "c")
			If ($smail -eq "c")
			{
				"Cancelled"
				Return
			}
			ElseIf ($smail -eq "n")
			{ $SendEmail = $false }
			ElseIf ($smail -eq "y")
			{
				
				Do
				{
					$MailFrom = Read-Host "Senders e-mail address. (""C"" to cancel)"
				}
				until
				($MailFrom -ne "" -or $MailFrom -eq "c")
				If ($MailFrom -eq "c")
				{
					"Cancelled"
					Return
				}
				Do
				{
					$MailTo = Read-Host "Recipients e-mail address (""C"" to cancel)"
				}
				until
				($MailTo -ne "" -or $MailTo -eq "c")
				If ($MailTo -eq "c")
				{
					"Cancelled"
					Return
				}
				Do
				{
					$MailServer = Read-Host "E-Mail server (""C"" to cancel)"
				}
				until
				($MailServer -ne "" -or $MailServer -eq "c")
				If ($MailServer -eq "c")
				{
					"Cancelled"
					Return
				}
				$SendEmail = $true
				
			}
			#endregion
			
			
			#...................................
			# Variables
			#...................................
			
			$now = Get-Date #Used for timestamps
			$date = $now.ToShortDateString() #Short date format for email message subject
			
			$report = @()
			
			$stats = @("DeviceID",
				"DeviceAccessState",
				"DeviceAccessStateReason",
				"DeviceModel"
				"DeviceType",
				"DeviceFriendlyName",
				"DeviceOS",
				"LastSyncAttemptTime",
				"LastSuccessSync"
			)
			
			$reportemailsubject = "Exchange ActiveSync Device Report - $date"
			$myDir = Get-ScriptDirectory
			$reportfile = "$myDir\ExchangeActiveSyncDeviceReport.csv"
			
			
			#...................................
			# Email Settings
			#...................................
			Write-host 'Enter your SMTP settings:'
			$MailTo = Read-host 'Enter receive Mail address'
			$MailFrom = Read-host 'Enter sender Mail address'
			$reportemailsubject = Read-host 'Enter Mail subject'
			$MailServer = Read-host 'Enter SMTP Mail server'
			
			$smtpsettings = @{
				To		   = $MailTo
				From	   = $MailFrom
				Subject    = $reportemailsubject
				SmtpServer = $MailServer
			}
			
			
			#...................................
			# Initialize
			#...................................
			
			#Add Exchange 2010/2013 snapin if not already loaded in the PowerShell session
			if (!(Get-PSSnapin | where { $_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010" }))
			{
				try
				{
					Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
				}
				catch
				{
					#Snapin was not loaded
					Write-Warning $_.Exception.Message
					EXIT
				}
				. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
				Connect-ExchangeServer -auto -AllowClobber
			}
			
			
			#...................................
			# Script
			#...................................
			
			Write-Host "Fetching list of mailboxes with EAS device partnerships"
			
			$MailboxesWithEASDevices = @(Get-CASMailbox -Resultsize Unlimited | Where { $_.HasActiveSyncDevicePartnership })
			
			Write-Host "$($MailboxesWithEASDevices.count) mailboxes with EAS device partnerships"
			
			Foreach ($Mailbox in $MailboxesWithEASDevices)
			{
				
				$EASDeviceStats = @(Get-ActiveSyncDeviceStatistics -Mailbox $Mailbox.Identity -WarningAction SilentlyContinue)
				
				Write-Host "$($Mailbox.Identity) has $($EASDeviceStats.Count) device(s)"
				
				$MailboxInfo = Get-Mailbox $Mailbox.Identity | Select DisplayName, PrimarySMTPAddress, OrganizationalUnit
				
				Foreach ($EASDevice in $EASDeviceStats)
				{
					Write-Host -ForegroundColor Green "Processing $($EASDevice.DeviceID)"
					
					$lastsyncattempt = ($EASDevice.LastSyncAttemptTime)
					
					if ($lastsyncattempt -eq $null)
					{
						$syncAge = "Never"
					}
					else
					{
						$syncAge = ($now - $lastsyncattempt).Days
					}
					
					#Add to report if last sync attempt greater than Age specified
					if ($syncAge -ge $Age -or $syncAge -eq "Never")
					{
						Write-Host -ForegroundColor Yellow "$($EASDevice.DeviceID) sync age of $syncAge days is greater than $age, adding to report"
						
						$reportObj = New-Object PSObject
						$reportObj | Add-Member NoteProperty -Name "Display Name" -Value $MailboxInfo.DisplayName
						$reportObj | Add-Member NoteProperty -Name "Organizational Unit" -Value $MailboxInfo.OrganizationalUnit
						$reportObj | Add-Member NoteProperty -Name "Email Address" -Value $MailboxInfo.PrimarySMTPAddress
						$reportObj | Add-Member NoteProperty -Name "Sync Age (Days)" -Value $syncAge
						
						Foreach ($stat in $stats)
						{
							$reportObj | Add-Member NoteProperty -Name $stat -Value $EASDevice.$stat
						}
						
						$report += $reportObj
					}
				}
			}
			
			Write-Host -ForegroundColor White "Saving report to $reportfile"
			$report | Export-Csv -NoTypeInformation $reportfile -Encoding UTF8
			
			
			if ($SendEmail)
			{
				
				$reporthtml = $report | ConvertTo-Html -Fragment
				
				$htmlhead = "<html>
                           <style>
                           BODY{font-family: Arial; font-size: 8pt;}
                           H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
                           H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
                           H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
                           TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
                           TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
                           TD{border: 1px solid #969595; padding: 5px; }
                           td.pass{background: #B7EB83;}
                           td.warn{background: #FFF275;}
                           td.fail{background: #FF2626; color: #ffffff;}
                           td.info{background: #85D4FF;}
                           </style>
                           <body>
                <p>Report of Exchange ActiveSync device associations with greater than $age days since last sync attempt as of $date. CSV version of report attached to this email.</p>"
				
				$htmltail = "</body></html>"
				
				$htmlreport = $htmlhead + $reporthtml + $htmltail
				
				Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -Attachments $reportfile
				
				$Reboot = $false
				
			}
		}
		
		function O365ExportLastLogonDate($O365Creds)
		{
			#      Export Office 365 User Last Logon Date to CSV File
			################################################################################################################################################################
			# Script accepts 3 parameters from the command line
			#
			# Office365Username - Mandatory - Administrator login ID for the tenant we are querying
			# Office365Password - Mandatory - Administrator login password for the tenant we are querying
			# UserIDFile - Optional - Path and File name of file full of UserPrincipalNames we want the Last Logon Dates for.  Seperated by New Line, no header.
			#
			#
			# To run the script
			#
			# .\Get-LastLogonStats.ps1 -Office365Username admin@xxxxxx.onmicrosoft.com -Office365Password Password123 -InputFile c:\Files\InputFile.txt
			#
			# NOTE: If you do not pass an input file to the script, it will return the last logon time of ALL mailboxes in the tenant.  Not advisable for tenants with large
			# user count (< 3,000) 
			#
			# Author:                        Alan Byrne
			# Version:                       1.0
			# Last Modified Date:     16/08/2012
			# Last Modified By:       Alan Byrne
			################################################################################################################################################################
			
			#Ask for O365 credentials
			if ($O365Creds -eq $null)
			{
				$O365Creds = Get-Credential -Message "Enter your O365 credentials"
				
			}
			$userIDFile = Read-Host "Enter path of file with list of UPNs. If Empty, all users are included."
			
			#Constant Variables
			$OutputFile = "LastLogonDate.csv" #The CSV Output file that is created, change for your purposes
			
			
			#Main
			Function O365ExportlastLogonMain
			{
				
				#Remove all existing Powershell sessions
				Get-PSSession | Remove-PSSession
				
				#Call ConnectTo-ExchangeOnline function with correct credentials
				$sessionID = ConnectTo-ExchangeOnline -O365Credentials $O365Creds
				
				#Prepare Output file with headers
				Out-File -FilePath $OutputFile -InputObject "UserPrincipalName,LastLogonDate" -Encoding UTF8
				
				#Check if we have been passed an input file path
				if ($userIDFile -ne "")
				{
					#We have an input file, read it into memory
					$objUsers = import-csv -Header "UserPrincipalName" $UserIDFile
				}
				else
				{
					#No input file found, gather all mailboxes from Office 365
					$objUsers = get-mailbox -ResultSize Unlimited | select UserPrincipalName
				}
				
				#Iterate through all users 
				Foreach ($objUser in $objUsers)
				{
					#Connect to the users mailbox
					$objUserMailbox = get-mailboxstatistics -Identity $($objUser.UserPrincipalName) | Select LastLogonTime
					
					#Prepare UserPrincipalName variable
					$strUserPrincipalName = $objUser.UserPrincipalName
					
					#Check if they have a last logon time. Users who have never logged in do not have this property
					if ($objUserMailbox.LastLogonTime -eq $null)
					{
						#Never logged in, update Last Logon Variable
						$strLastLogonTime = "Never Logged In"
					}
					else
					{
						#Update last logon variable with data from Office 365
						$strLastLogonTime = $objUserMailbox.LastLogonTime
					}
					
					#Output result to screen for debuging (Uncomment to use)
					#write-host "$strUserPrincipalName : $strLastLogonTime"
					
					#Prepare the user details in CSV format for writing to file
					$strUserDetails = "$strUserPrincipalName,$strLastLogonTime"
					
					#Append the data to file
					Out-File -FilePath $OutputFile -InputObject $strUserDetails -Encoding UTF8 -append
					"Result has been exported to $OutputFile in the current directory"
				}
				
				#Clean up session
				Remove-PSSession -Id $sessionID
			}
			. O365ExportlastLogonMain
		}
		
		function O365ListDistGroupsAndMemberships ($O365Creds)
		{
			#List all Distribution Groups and their Membership in Office 365
			################################################################################################################################################################ 
			# Origin of function:
			# 
			# Original file name: Get-DistributionGroupMembers.ps1
			# Author:                 Alan Byrne 
			# Version:                 2.0 
			# Last Modified Date:     16/08/2014 
			# Last Modified By:     Alan Byrne alan@cogmotive.com 	
			#
			# Modified for Exchange Suite and implemented as function
			# Last Modified Date: 13/12/2018
			# Last Modified By: dominic.manning@avectris.ch
			################################################################################################################################################################ 
			
			#Ask for O365 credentials
			if ($O365Creds -eq $null)
			{
				$O365Creds = Get-Credential -Message "Enter your O365 credentials"
				
			}
			#Constant Variables 
			$OutputFile = "DistributionGroupMembers.csv" #The CSV Output file that is created, change for your purposes 
			$arrDLMembers = @{ }
			
			#Remove all existing Powershell sessions 
			Get-PSSession | Remove-PSSession
			
			#Create remote Powershell session 
			$sessionID = ConnectTo-ExchangeOnline -O365Credentials $O365Creds
			
			#Prepare Output file with headers 
			Out-File -FilePath $OutputFile -InputObject "Distribution Group DisplayName,Distribution Group Email,Member DisplayName, Member Email, Member Type" -Encoding UTF8
			
			#Get all Distribution Groups from Office 365 
			$objDistributionGroups = Get-DistributionGroup -ResultSize Unlimited
			
			#Iterate through all groups, one at a time     
			Foreach ($objDistributionGroup in $objDistributionGroups)
			{
				
				write-host "Processing $($objDistributionGroup.DisplayName)..."
				
				#Get members of this group 
				$objDGMembers = Get-DistributionGroupMember -Identity $($objDistributionGroup.PrimarySmtpAddress)
				
				write-host "Found $($objDGMembers.Count) members..."
				
				#Iterate through each member 
				Foreach ($objMember in $objDGMembers)
				{
					Out-File -FilePath $OutputFile -InputObject "$($objDistributionGroup.DisplayName),$($objDistributionGroup.PrimarySMTPAddress),$($objMember.DisplayName),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)" -Encoding UTF8 -append
					write-host "`t$($objDistributionGroup.DisplayName),$($objDistributionGroup.PrimarySMTPAddress),$($objMember.DisplayName),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)"
				}
			}
			
			#Clean up session 
			Remove-PSSession -Id $sessionID
			
			$Reboot = $false
			
		}
		
		function O365MailTrafficStatsbyUser ($O365Creds)
		{
			#Office 365 Mail Traffic Statistics by User
			################################################################################################################################################################ 
			# Origin of function:
			# 
			# This script connects to Office 365 and retrieves detailed SMTP mail traffic statistics by user 
			# Requires Office 365 Wave 15 
			# 
			# Office365Username - Mandatory - Administrator login ID for the tenant we are querying 
			# Office365Password - Mandatory - Administrator login password for the tenant we are querying 
			#  
			# This script outputs the results to a CSV file called DetailedMessageStats.csv 
			# 
			# To run the script 
			# 
			# .\Get-DetailedMessageStats.ps1 -Office365Username admin@xxxxxx.onmicrosoft.com -Office365Password Password123  
			# 
			# 
			# Author:                 Alan Byrne 
			# Version:                 1.1
			# Last Modified Date:     17/01/2013 
			# Last Modified By:     Alan Byrne @alanmbyrne 
			#
			# Modified for Exchange Suite and implemented as function
			# Last Modified Date: 13/12/2018
			# Last Modified By: dominic.manning@avectris.ch
			################################################################################################################################################################ 
			
			#Ask for O365 credentials
			if ($O365Creds -eq $null)
			{
				$O365Creds = Get-Credential -Message "Enter your O365 credentials"
				
			}
			
			
			$OutputFile = "DetailedMessageStats.csv"
			
			$sessionID = ConnectTo-ExchangeOnline -O365Credentials $O365Creds
			
			Write-Host "Collecting Recipients..."
			
			#Collect all recipients from Office 365 
			$Recipients = Get-Recipient * -ResultSize Unlimited | select PrimarySMTPAddress
			
			$MailTraffic = @{ }
			foreach ($Recipient in $Recipients)
			{
				$MailTraffic[$Recipient.PrimarySMTPAddress.ToLower()] = @{ }
			}
			$Recipients = $null
			
			#Collect Message Tracking Logs (These are broken into "pages" in Office 365 so we need to collect them all with a loop) 
			$Messages = $null
			$Page = 1
			do
			{
				Write-Host "Collecting Message Tracking - Page $Page..."
				$CurrMessages = Get-MessageTrace -PageSize 5000 -Page $Page | Select Received, SenderAddress, RecipientAddress, Size
				$Page++
				$Messages += $CurrMessages
			}
			until ($CurrMessages -eq $null)
			
			Remove-PSSession $session
			
			Write-Host "Crunching Results..."
			
			#Read each message tracking entry and add it to a hash table 
			foreach ($Message in $Messages)
			{
				if ($Message.SenderAddress -ne $null)
				{
					if ($MailTraffic.ContainsKey($Message.SenderAddress))
					{
						$MessageDate = Get-Date -Date $Message.Received -Format yyyy-MM-dd
						
						if ($MailTraffic[$Message.SenderAddress].ContainsKey($MessageDate))
						{
							$MailTraffic[$Message.SenderAddress][$MessageDate]['Outbound']++
							$MailTraffic[$Message.SenderAddress][$MessageDate]['OutboundSize'] += $Message.Size
						}
						else
						{
							$MailTraffic[$Message.SenderAddress][$MessageDate] = @{ }
							$MailTraffic[$Message.SenderAddress][$MessageDate]['Outbound'] = 1
							$MailTraffic[$Message.SenderAddress][$MessageDate]['Inbound'] = 0
							$MailTraffic[$Message.SenderAddress][$MessageDate]['InboundSize'] = 0
							$MailTraffic[$Message.SenderAddress][$MessageDate]['OutboundSize'] += $Message.Size
						}
						
					}
				}
				
				if ($Message.RecipientAddress -ne $null)
				{
					if ($MailTraffic.ContainsKey($Message.RecipientAddress))
					{
						$MessageDate = Get-Date -Date $Message.Received -Format yyyy-MM-dd
						
						if ($MailTraffic[$Message.RecipientAddress].ContainsKey($MessageDate))
						{
							$MailTraffic[$Message.RecipientAddress][$MessageDate]['Inbound']++
							$MailTraffic[$Message.RecipientAddress][$MessageDate]['InboundSize'] += $Message.Size
						}
						else
						{
							$MailTraffic[$Message.RecipientAddress][$MessageDate] = @{ }
							$MailTraffic[$Message.RecipientAddress][$MessageDate]['Inbound'] = 1
							$MailTraffic[$Message.RecipientAddress][$MessageDate]['Outbound'] = 0
							$MailTraffic[$Message.RecipientAddress][$MessageDate]['OutboundSize'] = 0
							$MailTraffic[$Message.RecipientAddress][$MessageDate]['InboundSize'] += $Message.Size
							
						}
					}
				}
			}
			
			Write-Host "Formatting Results..."
			
			#Build a table to format the results 
			$table = New-Object system.Data.DataTable "DetailedMessageStats"
			$col1 = New-Object system.Data.DataColumn Date, ([datetime])
			$table.columns.add($col1)
			$col2 = New-Object system.Data.DataColumn Recipient, ([string])
			$table.columns.add($col2)
			$col3 = New-Object system.Data.DataColumn Inbound, ([int])
			$table.columns.add($col3)
			$col4 = New-Object system.Data.DataColumn Outbound, ([int])
			$table.columns.add($col4)
			$col5 = New-Object system.Data.DataColumn InboundSize, ([int])
			$table.columns.add($col5)
			$col6 = New-Object system.Data.DataColumn OutboundSize, ([int])
			$table.columns.add($col6)
			
			#Transpose hashtable to datatable 
			ForEach ($Recipient in $MailTraffic.keys)
			{
				$RecipientName = $Recipient
				
				foreach ($Date in $MailTraffic[$RecipientName].keys)
				{
					$row = $table.NewRow()
					$row.Date = $Date
					$row.Recipient = $RecipientName
					$row.Inbound = $MailTraffic[$RecipientName][$Date].Inbound
					$row.Outbound = $MailTraffic[$RecipientName][$Date].Outbound
					$row.InboundSize = $MailTraffic[$RecipientName][$Date].InboundSize
					$row.OutboundSize = $MailTraffic[$RecipientName][$Date].OutboundSize
					$table.Rows.Add($row)
				}
			}
			
			#Export data to CSV and Screen 
			
			$table | sort Date, Recipient, Inbound, Outbound, InboundSize, OutboundSize | Out-GridView -Title "Messages Sent By User"
			
			$table | sort Date, Recipient, Inbound, Outbound, InboundSize, OutboundSize | export-csv $OutputFile
			
			Write-Host "Results saved to $OutputFile"
			$Reboot = $false
			#Clean up session 
			Remove-PSSession -Id $sessionID
		}
		
		function O365ExportLicenseReconcilation ($O365Creds)
		{
			# Export a Licence reconciliation report from Office 365
			# The Output will be written to this file in the current working directory
			cls
			$LogFile = "Office_365_Licenses.csv"
			"The Output will be written to $LogFile in the current working directory"
			sleep -Seconds 3
			#Import module MSOnline
			try
			{
				Import-Module MSOnline -ErrorAction Stop
			}
			catch
			{
				Throw "Module MSOnline could not be imported!"
				return
			}
			# Connect to Microsoft Online
			Connect-MsolService -Credential $O365Creds
			
			write-host "Connecting to Office 365..."
			
			# Get a list of all licences that exist within the tenant
			$licensetype = Get-MsolAccountSku | Where { $_.ConsumedUnits -ge 1 }
			
			# Loop through all licence types found in the tenant
			foreach ($license in $licensetype)
			{
				# Build and write the Header for the CSV file
				$headerstring = "DisplayName,UserPrincipalName,AccountSku"
				
				foreach ($row in $($license.ServiceStatus))
				{
					$headerstring = ($headerstring + "," + $row.ServicePlan.servicename)
				}
				
				Out-File -FilePath $LogFile -InputObject $headerstring -Encoding UTF8 -append
				
				write-host ("Gathering users with the following subscription: " + $license.accountskuid)
				
				# Gather users for this particular AccountSku
				$users = Get-MsolUser -all | where { $_.isLicensed -eq "True" -and $_.licenses.accountskuid -contains $license.accountskuid }
				
				# Loop through all users and write them to the CSV file
				foreach ($user in $users)
				{
					
					write-host ("Processing " + $user.displayname)
					
					$thislicense = $user.licenses | Where-Object { $_.accountskuid -eq $license.accountskuid }
					
					$datastring = ($user.displayname + "," + $user.userprincipalname + "," + $license.SkuPartNumber)
					
					foreach ($row in $($thislicense.servicestatus))
					{
						
						# Build data string
						$datastring = ($datastring + "," + $($row.provisioningstatus))
					}
					
					Out-File -FilePath $LogFile -InputObject $datastring -Encoding UTF8 -append
				}
				
				Out-File -FilePath $LogFile -InputObject " " -Encoding UTF8 -append
			}
			
			write-host ("Script Completed.  Results available in " + $LogFile)
			$Reboot = $false
		}
		
		function ExportMBXFolderPermissions
		{
			#Export mailbox permissions also from Office 365 to CSV file
			#Original script:
                           <#
			    .SYNOPSIS
			    Creates a HTML Report showing Sendas, Full Access and Send on Behalf Permission Information for Each Mailbox for your Exchange Organization, selected database or for a single user.
			   
			       Serkan Varoglu
			       
			       http:\\Mshowto.org
			       http:\\Get-Mailbox.org
			       
			       THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
			       RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
			       
			       Version 1.1, 5 March 2012 
				
				 -- Modified for Exchange Suite and implemented as function ---
					Last Modified Date: 13/12/2018
					Last Modified By: dominic.manning@avectris.ch
			       
			    .DESCRIPTION
			       
			    Creates a HTML Report showing Sendas, Full Access and Send on Behalf Permission Information for Each Mailbox for your Exchange Organization, selected database or for a single user.
			       By Default Inherited Send As permission and NT Authority\Self account will not be shown in the report unless you run the script with the parameters listed below.
			       Also by default all mailboxes will be reported if you want to report a single database, you can use -database parameter to specify your database name or you can get the report for a single user.
			       
			       .PARAMETER HTMLReport
			    Filename to write HTML Report to
			       
			       .PARAMETER Database
			    By default this script will report all mailboxes. If you want to report mailboxes in a single database, you can use this parameter to input your database name.
			       
			       .PARAMETER Mailbox
			    By default this script will report all mailboxes. If you want to report a single mailbox, you can use this parameter to input the mailbox you want to report.
			       
			       .SWITCH ShowInherited
			       If ShowInherited is added as switch the report will show Inherited Sendas permissions for mailboxes as well.
			       
			       .SWITCH ShowSelf
			       If ShowSelf is added as switch the report will show "NT Authority\Self" sendas permission for mailboxes as well.
			       
			       .EXAMPLE
			    Generate the HTML report 
			    .\Report-Permissions.ps1 -HTMLReport "C:\Users\SVaroglu\Desktop\MailboxPermissionReport.HTML"
			       
			#>
			
			param
			(
				[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $false, HelpMessage = 'File name to write HTML report to. For Example: c:\DistGroupReport.html')]
				[string]$HTMLReport,
				[Parameter(Position = 1, Mandatory = $false, ValueFromPipeline = $false, HelpMessage = 'This switch will list Inherited Sendas and Full Access permissions as well')]
				[switch]$ShowInherited,
				[Parameter(Position = 2, Mandatory = $false, ValueFromPipeline = $false, HelpMessage = 'This switch will list NT Authority\Self Permission as well')]
				[switch]$ShowSelf,
				[Parameter(Position = 3, Mandatory = $false, ValueFromPipeline = $false, HelpMessage = 'Choose a specific Database to Report')]
				$Database,
				[Parameter(Position = 4, Mandatory = $false, ValueFromPipeline = $false, HelpMessage = 'Choose a Mailbox to Report')]
				$Mailbox,
				[Parameter(Position = 5, Mandatory = $true, ValueFromPipeline = $false, HelpMessage = 'Specify your O365 credentials', ParameterSetName = "O365")]
				[switch]$O365,
				[Parameter(Position = 5, Mandatory = $true, ValueFromPipeline = $false, HelpMessage = 'Specify your O365 credentials', ParameterSetName = "O365")]
				$O365Creds
			)
			
			if ($O365)
			{
				$sessionID = ConnectTo-ExchangeOnline -O365Credentials $O365Creds
			}
			
			$Watch = [System.Diagnostics.Stopwatch]::StartNew()
			$WarningPreference = "SilentlyContinue"
			$ErrorActionPreference = "SilentlyContinue"
			$ShowInherited = $ShowInherited.IsPresent
			$ShowSelf = $ShowSelf.IsPresent
			$u = 1
			$s = 0
			$f = 0
			$b = 0
			$n = 0
			$nj = -1
			$gj = -1
			if (!$database) { $dbnull = 0 }
			if (!$mailbox) { $mbnull = 0 }
			if (($ShowSelf -like "false") -and ($ShowInherited -like "false"))
			{ $gentitle = "Mailboxes With Custom Permissions" }
			else
			{ $gentitle = "Mailboxes" }
			$gen = "<table border=""1"" bordercolor=""#4384D3"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#4384D3"" align=""center""><th colspan=""5""><font color=""#FFFFFF"">$($gentitle)</font></th></tr><tr>"
			$inh = "<table border=""1"" bordercolor=""#4384D3"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#4384D3"" align=""center""><th colspan=""5""><font color=""#FFFFFF"">Mailboxes With Only Inherited Permissions</font></th></tr><tr>"
			function _Progress
			{
				param ($PercentComplete,
					$Status)
				Write-Progress -id 1 -activity "Report for Mailboxes" -status $Status -percentComplete ($PercentComplete)
			}
			_Progress (($u * 100)/100) "Collecting Mailbox Information"
			if (!$database -and !$mailbox)
			{
				$mailboxes = get-mailbox -resultsize unlimited | Sort-Object name
			}
			elseif ($database -and !$mailbox)
			{
				$mailboxes = get-mailbox -database $database -resultsize unlimited | Sort-Object name
			}
			elseif (!$database -and $mailbox)
			{
				$mailboxes = get-mailbox $mailbox
			}
			else
			{
				Write-Host -ForegroundColor Cyan "Please choose database or single mailbox. Both Parameters can not be used at the same time. Ended without compiling a report."
				exit
			}
			$mcount = ($mailboxes | measure-object).count
			if ($mcount -eq 0)
			{
				Write-Host -ForegroundColor Cyan "No Mailbox Found. Ended without compiling a report. Please Check Your Input."
				exit
			}
			foreach ($mailbox in $mailboxes)
			{
				_Progress (($u * 95)/$mcount) "Processing $mailbox, $($u) of $($mcount) Mailboxes."
				$SenderBody = ""
				$FullBody = ""
				$BehalfBody = ""
				$sendbehalfs = Get-Mailbox $mailbox | select-object -expand grantsendonbehalfto | select-object -expand rdn | Sort-Object Unescapedname
				if (($ShowSelf -like "true") -and ($ShowInherited -like "true"))
				{
					$senders = Get-ADPermission $mailbox.identity | ?{ ($_.extendedrights -like "*send-as*") } | Sort-Object name
					$fullsenders = Get-Mailbox $mailbox | Get-MailboxPermission | ?{ ($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") }
				}
				elseif (($ShowSelf -like "false") -and ($ShowInherited -like "true"))
				{
					$senders = Get-ADPermission $mailbox.identity | ?{ ($_.extendedrights -like "*send-as*") -and ($_.User -notlike "NT Authority\self") } | Sort-Object name
					$fullsenders = Get-Mailbox $mailbox | Get-MailboxPermission | ?{ ($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\self*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") }
				}
				elseif (($ShowSelf -like "true") -and ($ShowInherited -like "false"))
				{
					$senders = Get-ADPermission $mailbox.identity | ?{ ($_.extendedrights -like "*send-as*") -and ($_.isinherited -like "false") } | Sort-Object name
					$fullsenders = Get-Mailbox $mailbox | Get-MailboxPermission | ?{ ($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") -and ($_.IsInherited -like "false") }
				}
				else
				{
					$senders = Get-ADPermission $mailbox.identity | ?{ ($_.extendedrights -like "*send-as*") -and ($_.isinherited -like "false") -and ($_.User -notlike "NT Authority\self") } | Sort-Object name
					$fullsenders = Get-Mailbox $mailbox | Get-MailboxPermission | ?{ ($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\self*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") -and ($_.IsInherited -like "false") }
				}
				if (!$senders -and !$fullsenders -and !$sendbehalfs)
				{
					$n++
					if ($nj -eq 4)
					{
						$inh += "</tr><tr><td>$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</td>"
						$nj = 0
					}
					else
					{
						$inh += "<td>$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</td>"
						$nj++
					}
				}
				else
				{
					if ($gj -eq 4)
					{
						$gen += "</tr><tr><td><a href=""#$($mailbox.Name)"">$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</a></td>"
						$gj = 0
					}
					else
					{
						$gen += "<td><a href=""#$($mailbox.Name)"">$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</a></td>"
						$gj++
					}
					$MailboxTable = "<table border=""1"" bordercolor=""#1F497B"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#1F497B"" align=""center""><th colspan=""3"" ><font color=""#FFFFFF""><a name=""$($mailbox.Name)"">$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</font></a></th></tr><tr>"
					if (!$senders)
					{
						$SenderBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#696969"" align=""center"" valign=""top""><font color=""#FFFF99"">No User has Send As Permission On This Mailbox</font></td></table></td>"
					}
					else
					{
						$SenderBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><tr><td colspan=""3"" align=""center"" valign=""top"" bgcolor=""#696969""><font color=""#FFFFE0"">Send-As Permissions</font></td></tr><tr bgcolor=""#878787"" align=""center"">
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Send as Permission Owner</font></td>
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Deny</font></td>
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Inherited</font></td>
                                        </tr>"
						foreach ($sender in $senders)
						{
							if (0, 2, 4, 6, 8 -contains "$sj"[-1] - 48)
							{
								$bgcolor = "'#E8E8E8'"
							}
							else
							{
								$bgcolor = "'#C8C8C8'"
							}
							$SenderBody += "<tr align=""center"" bgcolor=$($bgcolor)>"
							$SenderBody += "<td><font color=""#003333"">$($sender.user)</font></td>"
							if ($sender.deny -like "true") { $font = "red" }
							else { $font = "'#000000'" }
							$SenderBody += "<td><font color=$font>$($sender.deny)</font></td>"
							if ($sender.isinherited -like "false") { $font = "red" }
							else { $font = "'#000000'" }
							$SenderBody += "<td><font color=$font>$($sender.isinherited)</font></td>"
							$SenderBody += "</tr>"
							$sj++
						}
						$SenderBody += "</table></td>"
						$s++
					}
					
					if (!$fullsenders)
					{
						$FullBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#696969"" align=""center"" valign=""top""><font color=""#FFFF99"">No User has Full Access On This Mailbox</font></td></table></td>"
					}
					else
					{
						$FullBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><tr><td colspan=""3"" align=""center"" valign=""top"" bgcolor=""#696969""><font color=""#FFFFE0"">Full Access Permissions</font></td></tr><tr bgcolor=""#878787"" align=""center"">
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Full Access Permission Owner</font></td>
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Deny</font></td>
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Inherited</font></td>
                                        </tr>"
						foreach ($fullsender in $fullsenders)
						{
							if (0, 2, 4, 6, 8 -contains "$fj"[-1] - 48)
							{
								$bgcolor = "'#E8E8E8'"
							}
							else
							{
								$bgcolor = "'#C8C8C8'"
							}
							$FullBody += "<tr align=""center"" bgcolor=$($bgcolor)>"
							$FullBody += "<td><font color=""#003333"">$($fullsender.user)</font></td>"
							if ($fullsender.deny -like "true") { $font = "red" }
							else { $font = "'#000000'" }
							$FullBody += "<td><font color=$font>$($fullsender.deny)</font></td>"
							if ($fullsender.isinherited -like "false") { $font = "red" }
							else { $font = "'#000000'" }
							$FullBody += "<td><font color=$font>$($fullsender.isinherited)</font></td>"
							$FullBody += "</tr>"
							$fj++
						}
						$FullBody += "</table></td>"
						$f++
					}
					
					if (!$sendbehalfs)
					{
						$BehalfBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#696969"" align=""center"" valign=""top""><font color=""#FFFF99"">No User has Send on Behalf On This Mailbox</font></td></table></td>"
					}
					else
					{
						$BehalfBody += "<td align=""center"" valign=""top"" width=""33%"">
                                        <table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
                                        <tr><td align=""center valign=""top"" bgcolor=""#696969""><font color=""#FFFFE0"">Send on Behalf</font></td></tr>
                                        <tr><td bgcolor=""#878787"" nowrap=""nowrap""><font color=""#FFFFFF"">Send On Behalf Permission Owner</font></td></tr>"
						foreach ($sendbehalf in $sendbehalfs)
						{
							if (0, 2, 4, 6, 8 -contains "$bj"[-1] - 48)
							{
								$bgcolor = "'#E8E8E8'"
							}
							else
							{
								$bgcolor = "'#C8C8C8'"
							}
							$BehalfBody += "<tr align=""center"" bgcolor=$($bgcolor)>"
							$BehalfBody += "<td><font color=""#003333"">$($sendbehalf.unescapedname)</font></td>"
							$BehalfBody += "</tr>"
							$bj++
						}
						$BehalfBody += "</table></td>"
						$b++
					}
					$Table += $MailboxTable + $SenderBody + $FullBody + $BehalfBody + "</tr></table><br><a href=""#top"">&#9650;</a><hr /><br>"
				}
				$u++
			}
			_Progress (98) "Completing"
			if (($ShowSelf -like "false") -and ($ShowInherited -like "false"))
			{
				if (($dbnull -eq 0) -and ($mbnull -eq 0))
				{
					$Summary = "<table style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#FFE87C"" >In your Exchange Organization there are $($mcount) mailboxes present."
					$Summary += "Send as Permission explicity configured on $($s) of these mailboxes. Full Access Permission explicity configured on $($f) of these mailboxes. Send on Behalf explicity configured on $($b) of these mailboxes and $($n) mailbox has inherited permissions only.<td></table><br>"
				}
				elseif (($dbnull -ne 0) -and ($mbnull -eq 0))
				{
					$Summary = "<table style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#FFE87C"" >In $($database) mailbox database, there are $($mcount) mailboxes present."
					$Summary += "Send as Permission explicity configured on $($s) of these mailboxes. Full Access Permission explicity configured on $($f) of these mailboxes. Send on Behalf explicity configured on $($b) of these mailboxes and $($n) mailbox has inherited permissions only.<td></table><br>"
				}
				$Header = "
       <body>
       <font size=""1"" face=""Arial,sans-serif"">
       <h3 align=""center"">Mailbox Send As, Full Permission and Send on Behalf Report</h3>
       <h4 align=""center"">Generated $((Get-Date).ToString())</h4>"
				$inh += "</tr></table><br>"
				$gen += "</tr></table><br>"
				$Footer = "</table></center><br><br>
       Scripted by <a href=""http://www.get-mailbox.org"">Serkan Varoglu</a>.  
       Elapsed Time To Complete This Report: $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString())</body></html>"
				if (($dbnull -eq 0) -and ($mbnull -eq 0))
				{
					$Output = $Header + $Summary + $gen + $inh + "<br><hr /><br>" + $Table + $Footer
				}
				elseif (($dbnull -ne 0) -and ($mbnull -eq 0))
				{
					$Output = $Header + $Summary + $gen + $inh + "<br><hr /><br>" + $Table + $Footer
				}
				else
				{
					if (($s -eq 0) -and ($f -eq 0) -and ($b -eq 0))
					{
						$Note = "<center></font><b>Mailbox for $($Mailbox.name) ( $($Mailbox.primarysmtpaddress) ), does not have any explicit permissions set for Send As, Full Access or Send on Behalf</b></center>"
					}
					$Output = $Header + $Note + $Table + $Footer
				}
			}
			else
			{
				$Header = "
       <body>
       <font size=""1"" face=""Arial,sans-serif"">
       <h3 align=""center"">Mailbox Send As, Full Permission and Send on Behalf Report</h3>
       <a name=""top""><h4 align=""center"">Generated $((Get-Date).ToString())</h4></a>
       "
				$inh += "</tr></table><br>"
				$gen += "</tr></table><br>"
				$Footer = "</table></center><br><br>
       <font size=""1"" face=""Arial,sans-serif"">Scripted by <a href=""http://www.get-mailbox.org"">Serkan Varoglu</a>.  
       Elapsed Time To Complete This Report: $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString())</font></body></html>"
				$Output = $Header + $gen + $Table + $Footer
				
			}
			$Output | Out-File $HTMLReport
			#clean up session
			Remove-PSSession -Id $sessionID
			$Reboot = $false
			
		}
			
		function ConnectTo-ExchangeOnline
		{
			Param (
				[Parameter(
						   Mandatory = $true,
						   Position = 0)]
				[PSCredential]$O365Credentials
			)
			
			#Create remote Powershell session
			$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Credentials -Authentication Basic –AllowRedirection
			
			#Import the session
			Import-PSSession $Session -AllowClobber | Out-Null
			#return the session ID 
			Return $session.Id
		}
		
		
		function O365EXCWebServicePrerequisites
		{
			#Import Localized Data
			Import-LocalizedData -BindingVariable Messages
			#Load .NET Assembly for Windows PowerShell V2
			Add-Type -AssemblyName System.Core
			
			$webSvcInstallDirRegKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Exchange\Web Services\2.0" -PSProperty "Install Directory" -ErrorAction:SilentlyContinue
			if ($webSvcInstallDirRegKey -ne $null)
			{
				$moduleFilePath = $webSvcInstallDirRegKey.'Install Directory' + 'Microsoft.Exchange.WebServices.dll'
				Import-Module $moduleFilePath
			}
			else
			{
				$errorMsg = $Messages.InstallExWebSvcModule
				throw $errorMsg
				return
			}
		}
		
		Function New-OSCPSCustomErrorRecord
		{
			#This function is used to create a PowerShell ErrorRecord
			[CmdletBinding()]
			Param
			(
				[Parameter(Mandatory = $true, Position = 1)]
				[String]$ExceptionString,
				[Parameter(Mandatory = $true, Position = 2)]
				[String]$ErrorID,
				[Parameter(Mandatory = $true, Position = 3)]
				[System.Management.Automation.ErrorCategory]$ErrorCategory,
				[Parameter(Mandatory = $true, Position = 4)]
				[PSObject]$TargetObject
			)
			Process
			{
				$exception = New-Object System.Management.Automation.RuntimeException($ExceptionString)
				$customError = New-Object System.Management.Automation.ErrorRecord($exception, $ErrorID, $ErrorCategory, $TargetObject)
				return $customError
			}
		}
		
		Function Connect-OSCEXOWebService
		{
			#.EXTERNALHELP Connect-OSCEXOWebService-Help.xml
			
			[cmdletbinding()]
			Param
			(
				#Define parameters
				[Parameter(Mandatory = $true, Position = 1)]
				[System.Management.Automation.PSCredential]$Credential,
				[Parameter(Mandatory = $false, Position = 2)]
				[Microsoft.Exchange.WebServices.Data.ExchangeVersion]$ExchangeVersion = "Exchange2010_SP2",
				[Parameter(Mandatory = $false, Position = 3)]
				[string]$TimeZoneStandardName,
				[Parameter(Mandatory = $false)]
				[switch]$Force
			)
			Process
			{
				#Get specific time zone info
				if (-not [System.String]::IsNullOrEmpty($TimeZoneStandardName))
				{
					Try
					{
						$tzInfo = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneStandardName)
					}
					Catch
					{
						$PSCmdlet.ThrowTerminatingError($_)
					}
				}
				else
				{
					$tzInfo = $null
				}
				
				#Create the callback to validate the redirection URL.
				$validateRedirectionUrlCallback = {
					param ([string]$Url)
					if ($Url -eq "https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml")
					{
						return $true
					}
					else
					{
						return $false
					}
				}
				
				#Try to get exchange service object from global scope
				$existingExSvcVar = (Get-Variable -Name exService -Scope Global -ErrorAction:SilentlyContinue) -ne $null
				
				#Establish the connection to Exchange Web Service
				if ((-not $existingExSvcVar) -or $Force)
				{
					$verboseMsg = $Messages.EstablishConnection
					$PSCmdlet.WriteVerbose($verboseMsg)
					if ($tzInfo -ne $null)
					{
						$global:exService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(`
							[Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion, $tzInfo)
					}
					else
					{
						$global:exService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(`
							[Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion)
					}
					
					#Set network credential
					$userName = $Credential.UserName
					$exService.Credentials = $Credential.GetNetworkCredential()
					Try
					{
						#Set the URL by using Autodiscover
						$exService.AutodiscoverUrl($userName, $validateRedirectionUrlCallback)
						$verboseMsg = $Messages.SaveExWebSvcVariable
						$PSCmdlet.WriteVerbose($verboseMsg)
						Set-Variable -Name exService -Value $exService -Scope Global -Force
					}
					Catch [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverRemoteException]
					{
						$PSCmdlet.ThrowTerminatingError($_)
					}
					Catch
					{
						$PSCmdlet.ThrowTerminatingError($_)
					}
				}
				else
				{
					$verboseMsg = $Messages.FindExWebSvcVariable
					$verboseMsg = $verboseMsg -f $exService.Credentials.Credentials.UserName
					$PSCmdlet.WriteVerbose($verboseMsg)
				}
			}
		}
		
		Function Get-OSCEXOCalendarFolder
		{
			#.EXTERNALHELP Get-OSCEXOCalendarFolder-Help.xml
			
			[cmdletbinding(DefaultParameterSetName = "DisplayName")]
			Param
			(
				#Define parameters
				[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true, ParameterSetName = "DisplayName")]
				[string]$DisplayName,
				[Parameter(Mandatory = $true, Position = 1, ParameterSetName = "Path")]
				[string]$Path,
				[Parameter(Mandatory = $false, Position = 2)]
				[switch]$ExactMatch,
				[Parameter(Mandatory = $false, Position = 3)]
				[Microsoft.Exchange.WebServices.Data.FolderTraversal]$FolderTraversal = "Deep"
			)
			Begin
			{
				#Verify the existence of exchange service object
				if ($exService -eq $null)
				{
					$errorMsg = $Messages.RequireConnection
					$customError = New-OSCPSCustomErrorRecord `
															  -ExceptionString $errorMsg `
															  -ErrorCategory NotSpecified -ErrorID 1 -TargetObject $PSCmdlet
					$PSCmdlet.ThrowTerminatingError($customError)
				}
				
				#Add folder path to extended property for identifying folder that has same name
				$PR_FOLDER_PATHNAME = 26293
				$exPropDefPathName = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(`
					$PR_FOLDER_PATHNAME, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
				
				$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(`
					[Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
				
				#Define the view settings in a folder search operation.
				$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
				$folderView.Traversal = $FolderTraversal
				$folderView.PropertySet = $propertySet
				$folderView.PropertySet.Add($exPropDefPathName)
				
				#Bind Message Folder Root
				$rootCalFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(`
					$exService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
			}
			Process
			{
				Switch ($PSCmdlet.ParameterSetName)
				{
					"DisplayName" {
						#Prepare search filter to find folder with the specific display name
						if ($ExactMatch)
						{
							$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(`
								[Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $DisplayName)
						}
						else
						{
							$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring(`
								[Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $DisplayName)
						}
					}
					"Path" {
						#Prepare search filter to find folder that matches specific path
						$Path = $Path.TrimEnd("\")
						if ($ExactMatch)
						{
							$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(`
								$exPropDefPathName, $Path)
						}
						else
						{
							$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring(`
								$exPropDefPathName, $Path)
						}
					}
				}
				
				#Begin to find folders
				do
				{
					$findResults = $rootCalFolder.FindFolders($searchFilter, $folderView)
					foreach ($folder in $findResults.Folders)
					{
						$propertySet.Add($exPropDefPathName)
						$folderObject = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exService, $folder.Id, $propertySet)
						$PSCmdlet.WriteObject($folderObject)
					}
				}
				while ($findResults.MoreAvailable)
			}
			End { }
		}
		
		Function Set-OSCEXOCalendarFolderPermission
		{
			#.EXTERNALHELP Set-OSCEXOCalendarFolderPermission-Help.xml
			
			[cmdletbinding(SupportsShouldProcess = $true)]
			Param
			(
				#Define parameters
				[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)]
				[Microsoft.Exchange.WebServices.Data.Folder]$Folder,
				[Parameter(Mandatory = $true, Position = 2)]
				[string]$UserName,
				[Parameter(Mandatory = $true, Position = 3)]
				[Microsoft.Exchange.WebServices.Data.FolderPermissionLevel]$PermissionLevel
			)
			Begin
			{
				#Verify the existence of exchange service object
				if ($exService -eq $null)
				{
					$errorMsg = $Messages.RequireConnection
					$customError = New-OSCPSCustomErrorRecord `
															  -ExceptionString $errorMsg `
															  -ErrorCategory NotSpecified -ErrorID 1 -TargetObject $PSCmdlet
					$PSCmdlet.ThrowTerminatingError($customError)
				}
				
				#Verify user name
				$verifiedUserName = $exService.ResolveName($UserName,`
					[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly, $false)
				if ($verifiedUserName -ne $null)
				{
					$userSmtpAddress = $verifiedUserName[0].Mailbox.Address
				}
				else
				{
					$errorMsg = $Messages.CannotResolveUserName
					$errorMsg = $errorMsg -f $UserName
					$customError = New-OSCPSCustomErrorRecord `
															  -ExceptionString $errorMsg `
															  -ErrorCategory NotSpecified -ErrorID 1 -TargetObject $PSCmdlet
					$PSCmdlet.ThrowTerminatingError($customError)
				}
			}
			Process
			{
				Try
				{
					#Get old permission entry
					$oldPermission = $Folder.Permissions | Where-Object {`
						($_.UserId.PrimarySmtpAddress -eq $userSmtpAddress)
					}
					
					#Update the permission entry if it exists
					if ($oldPermission -ne $null)
					{
						if ($PSCmdlet.ShouldProcess($Folder.DisplayName))
						{
							$oldPermissionLevel = $oldPermission.PermissionLevel
							$oldPermission.PermissionLevel = $PermissionLevel
							$Folder.Update()
							$verboseMsg = $Messages.SucceededToUpdatePermision
							$verboseMsg = $verboseMsg -f $oldPermissionLevel, $PermissionLevel
							$PSCmdlet.WriteVerbose($verboseMsg)
						}
					}
					else
					{
						$warningMsg = $Messages.UserDoesNotExist
						$warningMsg = $warningMsg -f $userSmtpAddress
						$PSCmdlet.WriteWarning($warningMsg)
					}
				}
				Catch
				{
					$PSCmdlet.WriteError($_)
				}
			}
			End { }
		}
		
		Function Grant-OSCEXOCalendarFolderPermission
		{
			#.EXTERNALHELP Grant-OSCEXOCalendarFolderPermission-Help.xml
			
			[cmdletbinding(SupportsShouldProcess = $true)]
			Param
			(
				#Define parameters
				[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)]
				[Microsoft.Exchange.WebServices.Data.Folder]$Folder,
				[Parameter(Mandatory = $true, Position = 2)]
				[string]$UserName,
				[Parameter(Mandatory = $true, Position = 3)]
				[Microsoft.Exchange.WebServices.Data.FolderPermissionLevel]$PermissionLevel
			)
			Begin
			{
				#Verify the existence of exchange service object
				if ($exService -eq $null)
				{
					$errorMsg = $Messages.RequireConnection
					$customError = New-OSCPSCustomErrorRecord `
															  -ExceptionString $errorMsg `
															  -ErrorCategory NotSpecified -ErrorID 1 -TargetObject $PSCmdlet
					$PSCmdlet.ThrowTerminatingError($customError)
				}
				
				#Verify user name
				$verifiedUserName = $exService.ResolveName($UserName,`
					[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly, $false)
				if ($verifiedUserName -ne $null)
				{
					$userSmtpAddress = $verifiedUserName[0].Mailbox.Address
				}
				else
				{
					$errorMsg = $Messages.CannotResolveUserName
					$errorMsg = $errorMsg -f $UserName
					$customError = New-OSCPSCustomErrorRecord `
															  -ExceptionString $errorMsg `
															  -ErrorCategory NotSpecified -ErrorID 1 -TargetObject $PSCmdlet
					$PSCmdlet.ThrowTerminatingError($customError)
				}
			}
			Process
			{
				#Prepare new permission
				$userID = New-Object Microsoft.Exchange.WebServices.Data.UserId($userSmtpAddress)
				$newPermission = New-Object Microsoft.Exchange.WebServices.Data.FolderPermission($userID, $PermissionLevel)
				
				#Get old permission
				$oldPermission = $Folder.Permissions | Where-Object {`
					($_.UserId.PrimarySmtpAddress -eq $userSmtpAddress)
				}
				
				#Add permission if permission entry does not exist.
				if ($oldPermission -ne $null)
				{
					$warningMsg = $Messages.UserExists
					$warningMsg = $warningMsg -f $userSmtpAddress
					$PSCmdlet.WriteWarning($warningMsg)
				}
				else
				{
					if ($PSCmdlet.ShouldProcess($Folder.DisplayName))
					{
						Try
						{
							$Folder.Permissions.Add($newPermission) | Out-Null
							$Folder.Update()
							$verboseMsg = $Messages.SucceededToAddPermision
							$verboseMsg = $verboseMsg -f $userSmtpAddress, $PermissionLevel
							$PSCmdlet.WriteVerbose($verboseMsg)
						}
						Catch
						{
							$verboseMsg = $Messages.FailedToAddPermision
							$verboseMsg = $verboseMsg -f $userSmtpAddress, $PermissionLevel
							$PSCmdlet.WriteVerbose($verboseMsg)
							$PSCmdlet.WriteError($_)
						}
					}
				}
			}
			End { }
		}
		
		Function Revoke-OSCEXOCalendarFolderPermission
		{
			#.EXTERNALHELP Revoke-OSCEXOCalendarFolderPermission-Help.xml
			
			[cmdletbinding(SupportsShouldProcess = $true)]
			Param
			(
				#Define parameters
				[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)]
				[Microsoft.Exchange.WebServices.Data.Folder]$Folder,
				[Parameter(Mandatory = $true, Position = 2)]
				[string]$UserName,
				[Parameter(Mandatory = $true, Position = 3)]
				[Microsoft.Exchange.WebServices.Data.FolderPermissionLevel]$PermissionLevel
			)
			Begin
			{
				#Verify the existence of exchange service object
				if ($exService -eq $null)
				{
					$errorMsg = $Messages.RequireConnection
					$customError = New-OSCPSCustomErrorRecord `
															  -ExceptionString $errorMsg `
															  -ErrorCategory NotSpecified -ErrorID 1 -TargetObject $PSCmdlet
					$PSCmdlet.ThrowTerminatingError($customError)
				}
				
				#Verify user name
				$verifiedUserName = $exService.ResolveName($UserName,`
					[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly, $false)
				if ($verifiedUserName -ne $null)
				{
					$userSmtpAddress = $verifiedUserName[0].Mailbox.Address
				}
				else
				{
					$errorMsg = $Messages.CannotResolveUserName
					$errorMsg = $errorMsg -f $UserName
					$customError = New-OSCPSCustomErrorRecord `
															  -ExceptionString $errorMsg `
															  -ErrorCategory NotSpecified -ErrorID 1 -TargetObject $PSCmdlet
					$PSCmdlet.ThrowTerminatingError($customError)
				}
			}
			Process
			{
				#Get old permission entry
				$oldPermission = $Folder.Permissions | Where-Object {`
					($_.UserId.PrimarySmtpAddress -eq $userSmtpAddress) -and `
					($_.PermissionLevel -eq $PermissionLevel)
				}
				
				#Remove old permission entry if it exists
				if ($oldPermission -ne $null)
				{
					if ($PSCmdlet.ShouldProcess($Folder.DisplayName))
					{
						Try
						{
							$Folder.Permissions.Remove($oldPermission) | Out-Null
							$Folder.Update()
							$verboseMsg = $Messages.SucceededToRemovePermision
							$verboseMsg = $verboseMsg -f $userSmtpAddress, $PermissionLevel
							$PSCmdlet.WriteVerbose($verboseMsg)
						}
						Catch
						{
							$verboseMsg = $Messages.FailedToRemovePermision
							$verboseMsg = $verboseMsg -f $userSmtpAddress, $PermissionLevel
							$PSCmdlet.WriteVerbose($verboseMsg)
							$PSCmdlet.WriteError($_)
						}
					}
				}
				else
				{
					$warningMsg = $Messages.PermissionDoesNotExist
					$warningMsg = $warningMsg -f $userSmtpAddress, $PermissionLevel
					$PSCmdlet.WriteWarning($warningMsg)
				}
			}
			End { }
		}
		
		
		
		
		
		function Get-ScriptDirectory
		{
			<#
				.SYNOPSIS
					Get-ScriptDirectory returns the proper location of the script.
			
				.OUTPUTS
					System.String
				
				.NOTES
					Returns the correct path within a packaged executable.
			#>
			[OutputType([string])]
			param ()
			if ($null -ne $hostinvocation)
			{
				Split-Path $hostinvocation.MyCommand.path
			}
			else
			{
				Split-Path $script:MyInvocation.MyCommand.Path
			}
		}
		#endregion
		switch ($Choice)
		{
			#region Option  1) Windows Update
			1 {
				#      Windows Update
				Invoke-Expression "$env:windir\system32\wuapp.exe startmenu"
			}
			#endregion   
			#region Option  2) Mailbox Requirement Check
			2 {
				#      Mailbox Requirement Check
				Check-MBXprereq
			}
			#endregion   
			#region Option  3) Edge Transport Requirement Check
			3 {
				#      Edge Transport Requirement Check
				Check-EdgePrereq
			}
			#endregion   
			#region Option  4) Prep Mailbox Role - Part 1 - CU3+
			4 {
				#      Prep Mailbox Role - Part 1 - CU3+
				ModuleStatus -name ServerManager
				Install-WindowsFeature RSAT-ADDS
				Install-WindowsFeature NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
				HighPerformance
				PowerMgmt
				$Reboot = $true
			}
			#endregion   
			#region Option  5) Prep Mailbox Role - Part 2 - CU3+
			5 {
				#      Prep Mailbox Role - Part 2 - CU3+
				ModuleStatus -name ServerManager
				Install-WinUniComm4
				$Reboot = $true
			}
			#endregion   
			#region Option  6) Prep Exchange Transport - CU3+
			6 {
				#      Prep Exchange Transport - CU3+
				Install-windowsfeature ADLDS
			}
			#endregion   
			#region Option  7) Install -One-Off - Windows Features [MBX] - CU3+
			7 {
				#      Install -One-Off - Windows Features [MBX] - CU3+
				ModuleStatus -name ServerManager
				Install-WindowsFeature NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
			}
			#endregion   
			#region Option  8) Install - One Off - Unified Communications Managed API 4.0 - CU3+
			8 {
				#      Install - One Off - Unified Communications Managed API 4.0 - CU3+
				Install-WinUniComm4
			}
			#endregion   
			#region Option  9) Prepare Schema
			9 {
				#      Prepare Schema
				Write-Verbose -Message "Insert the installation Media or Mount the Exchange ISO!!!" -verbose
				$ISOpath = Read-Host -Prompt 'Enter the installation Path <e.g. e:>'
				Write-Verbose -Message "start Prepare Schema" -verbose
				cd $ISOpath
				.\Setup.exe /PrepareSchema /IAcceptExchangeServerLicenseTerms
				Write-Host "done" -ForegroundColor green
			}
			#endregion   
			#region Option 10) Prepare Active Directory and Domains
			10 {
				#      Prepare Active Directory and Domains
				Write-Verbose -Message "Insert the installation Media or Mount the Exchange ISO!!!" -verbose
				$ISOpath = Read-Host -Prompt 'Enter the installation Path <e.g. e:>'
				$domainorg = Read-Host -Prompt 'Set the name of your Exchange Organisation <e.g. Contoso>'
				$domainset = Read-Host -Prompt 'Set the name of your Domain <e.g. contoso.com>'
				Write-Verbose -Message "start Prepare Active Directory" -verbose
				cd $ISOpath
				.\Setup.exe /PrepareAD /OrganizationName: $domainorg /IAcceptExchangeServerLicenseTerms
				Write-Host "done" -ForegroundColor green
				Write-Verbose -Message "start Prepare Active Directory domains" -verbose
				cd $ISOpath
				.\Setup.exe /PrepareDomain:$domainset /IAcceptExchangeServerLicenseTerms
			}
			#endregion   
			#region Option 11) Set Power Plan to High Performance
			11 {
				#      Set Power Plan to High Performance
				highperformance
			}
			#endregion   
			#region Option 12) Disable Power Management for NICs
			12 {
				#      Disable Power Management for NICs
				PowerMgmt
			}
			#endregion   
			#region Option 13) Disable SSL 3.0 Support
			13 {
				#      Disable SSL 3.0 Support
				DisableSSL3
			}
			#endregion   
			#region Option 14) Disable RC4 Support
			14 {
				#      Disable RC4 Support
				DisableRC4
			}
			#endregion   
			#region Option 30) INSTALL EXCHANGE SERVER
			30 {
				#      INSTALL EXCHANGE SERVER
				Write-Verbose -Message "Insert the installation Media or Mount the Exchange ISO!!!" -verbose
				$ISOpath = Read-Host -Prompt 'Enter the installation Path <f.e. "e:">'
				Write-Verbose -Message "start Exchange Setup" -verbose
				cd $ISOpath
				.\setup /Mode:Install /Role:Mailbox /IAcceptExchangeServerLicenseTerms
			}
			#endregion   
			#region Option 40) Configure Page File
			40 {
				#   Add Windows Defender Exclusions
				ConfigurePageFile
			}
			#endregion   
			#region Option 41) Show Exchange URIs
			41 {
				
					#   Show Exchange URIs
				ShowEXCURI
				"waiting 10 seconds..."
				sleep -Seconds 10
			}
			#endregion   
			#region Option 42) Configure Exchange URLs
			42 {
				$server = (Get-ExchangeServer).fqdn
				$InternalURL = Read-Host "Please enter the internal URL. (Mandatory)"
				$ExternalURL = Read-Host "Please enter the external URL. (Mandatory)"
				$AutodiscoverSCP = Read-Host "Please enter the Autodiscover SCP URL. (Optional)"
				$SSLInt = Read-Host "SSL for internal Outlook Anywhere? [Y/N]"
				$SSLExt = Read-Host "SSL for external Outlook Anywhere? [Y/N]"
				
				if ($sslint -eq "y")
				{
					$InternalSSL = $true
				}
				Else
				{
					$InternalSSL = $false
				}
				if ($SSLExt -eq "y")
				{
					$ExternalSSL = $true
				}
				Else
				{
					$ExternalSSL = $false
				}
				
				ConfigureEXCURL($server, $InternalURL, $ExternalURL, $AutodiscoverSCP, $InternalSSL, $ExternalSSL)
			}
			#endregion   
			#region Option 43) Disable UAC
			43 {
				#   Disable UAC
				New-ItemProperty -Path HKLM:Software\Microsoft\Windows\CurrentVersion\policies\system -Name EnableLUA -PropertyType DWord -Value 0 -Force
			}
			#endregion   
			#region Option 44) Disable Windows Firewall
			44 {
				#   Disable Windows Firewall
				Write-Host 'Disabeling Windows Firewall...' -ForegroundColor Yellow
				Set-NetFirewallProfile -Profile Domain, Public, Private -Enabled False
				Write-Host 'Done' -ForegroundColor Green
			}
			#endregion   
			#region Option 45) Create receive connector
			45 {
				#   Create receive connector
				$rcServer = Read-Host 'Enter Exchange server hostname for which to create the receive connector <e.g. Exc-srv001>'
				$rc1 = Read-Host 'Set Name of the receive connector <e.g. "Inbound SMTP Mailgateway">'
				$RemoteIPR1 = Read-Host 'Set the first Remote IP address <e.g. 000.000.000.000>'
				$RemoteIPR2 = Read-Host 'Set the second Remote IP address <e.g. 000.000.000.000>'
				Write-Host 'Setting up Receive Connector for Exchange server: $rcServer...' -ForegroundColor White
				New-ReceiveConnector -Name $rc1 -Bindings ("0.0.0.0:25") -RemoteIPRanges '$RemoteIPR1', '$RemoteIPR2' -MaxMessageSize 30MB -TransportRole FrontendTransport -Usage Custom -Server $rcServer -AuthMechanism 'TLS' -PermissionGroups 'AnonymousUsers'
				Write-Host 'Done' -ForegroundColor Green
			}
			#endregion   
			#region Option 46) Create send connector
			46 {
				#   Create send connector
				$scServer = Read-Host 'Enter Exchange server hostname for which to create the send connector <e.g. Exc-srv001>'
				$sc1 = Read-Host 'Set Name of the send connector <e.g. "Outbound to Internet">'
				$sRemoteIPR1 = Read-Host 'Set the first Source Transport Server <e.g. Exc-srv001>'
				$sRemoteIPR2 = Read-Host 'Set the second Source Transport Server <e.g. Exc-srv002>'
				$sAddressSpace = Read-Host 'Set the address space: <e.g. SMTP:*.contoso.com>'
				$sSmartHost = Read-Host 'Set Smarthost <e.g. 000.000.000.000 or SM.contoso.com>'
				Write-Host 'Setting up Send Connector for Exchange server: $scServer...' -ForegroundColor White
				New-SendConnector -Name $sc1 -AddressSpaces $sAddressSpace -SourceTransportServers '$sRemoteIPR1', '$sRemoteIPR2' -FrontendProxyEnabled:$false -SmartHosts $sSmartHost
				Write-Host 'Done' -ForegroundColor Green
			}
			#endregion   
			#region Option 47) Create DAG
			47 {
				#   Create DAG
				Write-Host "Starting setup to create DAG" -ForegroundColor White
				$DAGName = Read-Host 'Enter Name for DAG <e.g. DAG01>'
				$Witness = Read-Host 'Enter Hostname of Witness server <e.g. ABG-SRV01>'
				$WitnessPath = Read-Host 'Enter the local Path form Witness server where the Directory will be located: <e.g. C:\FSW\VMBDNDAGEKZ01>'
				Write-Host -Verbose 'Be sure that the Exchange permissions on the Witness server are set correctly!' -ForegroundColor Yellow -BackgroundColor Black
				$EXC01 = Read-Host 'Enter Hostname of the first Exchange server <e.g. SRV-EX01>'
				$EXC02 = Read-Host 'Enter Hostname of the second Exchange server <e.g. SRV-EX02>'
				New-DatabaseAvailabilityGroup -Name $DAGName -WitnessServer $Witness -WitnessDirectory $WitnessPath
				Add-DatabaseAvailabilityGroupServer -Identity $DAGName -MailboxServer $EXC01
				Add-DatabaseAvailabilityGroupServer -Identity $DAGName -MailboxServer $EXC02
				Get-DatabaseAvailabilityGroup $DAGName -Status
				Get-DatabaseAvailabilityGroup $DAGName -Status | fl *witness*
				Write-Host 'Done' -ForegroundColor Green
			}
			#endregion   
			#region Option 48) Create Exchange Hybrid mode
			48 {
				"This function is not yet implemented"
				#   Create Exchange Hybrid mode
				
			}
			#endregion   
			#region Option 49) Create Certificate request
			49 {
				#   Create Certificate request
				####################
				# Prerequisite check
				####################
				if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
				{
					Write-Host "Administrator priviliges are required. Please restart this script with elevated rights." -ForegroundColor Red
					Pause
					Throw "Administrator priviliges are required. Please restart this script with elevated rights."
				}
				
				
				#######################
				# Setting the variables
				#######################
				$UID = [guid]::NewGuid()
				$files = @{ }
				$files['settings'] = "$($env:TEMP)\$($UID)-settings.inf";
				$files['csr'] = "$($env:TEMP)\$($UID)-csr.req"
				
				
				$request = @{ }
				$request['SAN'] = @{ }
				
				Write-Host "keep it simple but significant" -ForegroundColor green
				Write-Host "Enter the Certificate informations below" -ForegroundColor cyan
				$request['CN'] = Read-Host "Common Name (e.g. company.com)"
				$request['O'] = Read-Host "Organisation (e.g. Company Ltd)"
				$request['OU'] = Read-Host "Organisational Unit (e.g. IT)"
				$request['L'] = Read-Host "City (e.g. Amsterdam)"
				$request['S'] = Read-Host "State (e.g. Noord-Holland)"
				$request['C'] = Read-Host "Country (e.g. NL)"
				
				###########################
				# Subject Alternative Names
				###########################
				$i = 0
				Do
				{
					$i++
					$request['SAN'][$i] = read-host "Subject Alternative Name $i (e.g. alt.company.com / leave empty for none)"
					if ($request['SAN'][$i] -eq "")
					{
						
					}
					
				}
				until ($request['SAN'][$i] -eq "")
				
				# Remove the last in the array (which is empty)
				$request['SAN'].Remove($request['SAN'].Count)
				
				#########################
				# Create the settings.inf
				#########################
				$settingsInf = "
                           [Version] 
                           Signature=`"`$Windows NT`$ 
                           [NewRequest] 
                           KeyLength =  2048
                           Exportable = TRUE 
                           MachineKeySet = TRUE 
                           SMIME = FALSE
                           RequestType =  PKCS10 
                           ProviderName = `"Microsoft RSA SChannel Cryptographic Provider`" 
                           ProviderType =  12
                           HashAlgorithm = sha256
                           ;Variables
                           Subject = `"CN={{CN}},OU={{OU}},O={{O}},L={{L}},S={{S}},C={{C}}`"
                           [Extensions]
                           {{SAN}}


                           ;Certreq info
                           ;http://technet.microsoft.com/en-us/library/dn296456.aspx
                           ;CSR Decoder
                           ;https://certlogik.com/decoder/
                           ;https://ssltools.websecurity.symantec.com/checker/views/csrCheck.jsp
                           "
				
				$request['SAN_string'] = & {
					if ($request['SAN'].Count -gt 0)
					{
						$san = "2.5.29.17 = `"{text}`"
"
						Foreach ($sanItem In $request['SAN'].Values)
						{
							$san += "_continue_ = `"dns=" + $sanItem + "&`"
"
						}
						return $san
					}
				}
				
				$settingsInf = $settingsInf.Replace("{{CN}}", $request['CN']).Replace("{{O}}", $request['O']).Replace("{{OU}}", $request['OU']).Replace("{{L}}", $request['L']).Replace("{{S}}", $request['S']).Replace("{{C}}", $request['C']).Replace("{{SAN}}", $request['SAN_string'])
				
				# Save settings to file in temp
				$settingsInf > $files['settings']
				
				# Done, we can start with the CSR
				Clear-Host
				
				#################################
				# CSR TIME
				#################################
				
				# Display summary
				Write-Host "Certificate information
                           Common name: $($request['CN'])
                           Organisation: $($request['O'])
                           Organisational unit: $($request['OU'])
                           City: $($request['L'])
                           State: $($request['S'])
                           Country: $($request['C'])

                           Subject alternative name(s): $($request['SAN'].Values -join ", ")

                           Signature algorithm: SHA256
                           Key algorithm: RSA
                           Key size: 2048

" -ForegroundColor Yellow
				
				certreq -new $files['settings'] $files['csr'] > $null
				
				# Output the CSR
				$CSR = Get-Content $files['csr']
				Write-Output $CSR
				Write-Host "
"
				
				# Set the Clipboard (Optional)
				Write-Host "Copy CSR to clipboard? (y|n): " -ForegroundColor Yellow -NoNewline
				if ((Read-Host) -ieq "y")
				{
					$csr | clip
					Write-Host "Check your ctrl+v
"
				}
				
				
				########################
				# Remove temporary files
				########################
				$files.Values | ForEach-Object {
					Remove-Item $_ -ErrorAction SilentlyContinue
				}
			}
			#endregion   
			#region Option 50) Set mailaddress policies
			50 {
				#   set mailaddress policies
				Write-Host "Enter your Domains to create the Mail address Policies" -ForegroundColor White
				$Name1 = Read-Host "Enter the Name you wanna use for the internal Policy e.g. fabrikam-local"
				$Dom1 = Read-Host "Enter your internal Exchange Domain e.g. fabrikam.local"
				$Name2 = Read-Host "Enter the Name you wanna use for primary external Domain Policy e.g. fabrikam-extern"
				$Dom2 = Read-Host "Enter your primary external Exchange Domain e.g. fabrikam.com"
				$LocPat2 = Read-Host "Enter the Member Group OU e.g. OU=FABRIKAM,OU=CUSTOMERS,DC=fabrikam,DC=local"
				$Name3 = Read-Host "Enter the Name you wanna use for primary Accepted Domain Policy e.g. contoso-extern"
				$Dom3 = Read-Host "Enter your first Accepted Exchange Domain e.g. contoso.com"
				$LocPat3 = Read-Host "Enter the Member Group OU e.g. OU=CONTOSO,OU=CUSTOMERS,DC=fabrikam,DC=local"
				$Name4 = Read-Host "Enter the Name you wanna use for secondary Accepted Domain Policy e.g. abstergo-extern"
				$Dom4 = Read-Host "Enter your second Accepted Exchange Domain e.g. abstergo.ch"
				$LocPat4 = Read-Host "Enter the Member Group OU e.g. OU=ABSTERGO,OU=CUSTOMERS,DC=fabrikam,DC=local"
				# Create Mailaddress Policy for Resources
				Write-Host "Creating the 1st Mailaddress Policy $Name1 for Resources..." -ForegroundColor cyan
				New-EmailAddressPolicy -Name $Name1 -EnabledPrimarySMTPAddressTemplate 'SMTP:alias@$Dom1' -IncludedRecipients 'Resources' -Priority 1
				Write-Host "Done!" -ForegroundColor green
				
				# Create primary Mailaddress Policy
				Write-Host "Creating the 2nd Mailaddress Policy $Name2 for the Domain $Dom2 ..." -ForegroundColor cyan
				New-EmailAddressPolicy -Name $Name2 -EnabledPrimarySMTPAddressTemplate 'SMTP:%g.%i.%s@$Dom2' -RecipientFilter { ((MemberOfGroup -eq $LocPat2) -and (RecipientType -eq 'UserMailbox')) } -Priority 2
				Set-EmailAddressPolicy $Name2 -EnabledEmailAddressTemplates SMTP:%g.%i.%s@$Dom2, smtp:%g.%i.%s@$Dom1, smtp:%1g.%s@$Dom1, smtp:alias@$Dom1
				Write-Host "Done!" -ForegroundColor green
				
				# Create first Accepted Domain Mailaddress Policy
				Write-Host "Creating the 3rd Mailaddress Policy $Name3 for the Accepted Domain $Dom3 ..." -ForegroundColor cyan
				New-EmailAddressPolicy -Name $Name3 -EnabledPrimarySMTPAddressTemplate 'SMTP:%g.%i.%s@$Dom3' -RecipientFilter { ((MemberOfGroup -eq $LocPat3) -and (RecipientType -eq 'UserMailbox')) } -Priority 3
				Set-EmailAddressPolicy $Name3 -EnabledEmailAddressTemplates SMTP:%g.%i.%s@$Dom3, smtp:%g.%i.%s@$Dom1, smtp:%1g.%s@$Dom1, smtp:alias@$Dom1
				Write-Host "Done!" -ForegroundColor green
				
				# Create first Accepted Domain Mailaddress Policy
				Write-Host "Creating the 4th Mailaddress Policy $Name4 for the Accepted Domain $Dom4 ..." -ForegroundColor cyan
				New-EmailAddressPolicy -Name $Name4 -EnabledPrimarySMTPAddressTemplate 'SMTP:%g.%i.%s@certum.ch' -RecipientFilter { ((MemberOfGroup -eq $LocPat4) -and (RecipientType -eq 'UserMailbox')) } -Priority 4
				Set-EmailAddressPolicy $Name4 -EnabledEmailAddressTemplates SMTP:%g.%i.%s@$Dom4, smtp:%g.%i.%s@$Dom1, smtp:%1g.%s@$Dom1, smtp:alias@$Dom1
				Write-Host "Done!" -ForegroundColor green
				
				# Enable all Policies
				Write-Host "Enable all Address Policies..." -ForegroundColor cyan
				Get-EmailAddressPolicy $Name1 | Update-EmailAddressPolicy
				Get-EmailAddressPolicy $Name2 | Update-EmailAddressPolicy
				Get-EmailAddressPolicy $Name3 | Update-EmailAddressPolicy
				Get-EmailAddressPolicy $Name4 | Update-EmailAddressPolicy
				Write-Host "Done!" -ForegroundColor green
				
				# Information
				Write-Host "All Mailaddress Policies are created! See the Summary below..." -ForegroundColor magenta -Verbose
				Get-EmailAddressPolicy
			}
			#endregion   
			#region Option 51) Enable UM for all Mailboxes
			51 {
				#   Enable UM for all Mailboxes
				#region log file
				$date = (Get-Date -Format yyyyMMdd_HHmm) #create time stamp
				$log = "$PSScriptRoot\$date-EnableUM.Log" #define path and name, incl. time stamp for log file
				#endregion
				
				Write-Host '--- keep it simple, but significant ---' -ForegroundColor magenta
				
				#region environment selection, modules and credentials
				#show options for environment selection
				$ExcOpt = Read-Host "Choose environment to connect to. 
                           [1] O365 
                           [2] On-Premises
                           Your option"
				
				#get credentials according to the selected environment
				switch ($ExcOpt)
				{
					1  {
						#If 1 is selected
						
						try
						{
							# Check if AzureAD module is available and import it. 
							if (!(Get-Module AzureAD) -or !(Get-Module AzureADPreview))
							{
								Import-Module AzureAD -ErrorAction Stop
								$AADModule = 'AAD'
								(Get-Date -Format G) + " " + "Azure AD module loaded" | Tee-Object -FilePath $log -Append
								
							}
							
						}
						
						catch
						{
							
							cls
							try
							{
								#Try to load MSonline module, if Azure AD module is not available
								Import-Module MSOnline -ErrorAction Stop
								$AADModule = 'MSOnline'
								(Get-Date -Format G) + " " + "MSOnline module loaded" | Tee-Object -FilePath $log -Append
							}
							catch
							{
								#If no module is available, show option to open download page for MSonline module
								Write-Host "For O365 environments you first need to install MSOnline, AzureAD, or AzureADPreview module!" -ForegroundColor Red
								Write-Host "Please install one of the modules and restart the script." -ForegroundColor Cyan
								""
								
								$red = Read-Host "Do you want to be redirected to the MS download page for the MSOnline module? [Y] Yes, [N] No. Default is No."
								switch ($red)
								{
									Y { [system.Diagnostics.Process]::Start('http://connect.microsoft.com/site1164/Downloads/DownloadDetails.aspx?DownloadID=59185') }
									N { "Script will end now." }
									default { "Script will end now." }
								}
								return
							}
							
							
							
						}
						#Ask for O365 credentials
						"O365 selected"; $O365Creds = Get-Credential -Message 'Enter your O365 credentials'
						
					}
					2  {
						#If 2 is selected, ask for On-Prem credentials
						"On-Prem selected"; $OnPremCreds = Get-Credential -Message 'Enter your Exchange On-Prem credentials'
					}
					default { Write-Host "Please enter 1, or 2" -ForegroundColor Red; return }
				}
				#endregion
				
				#region select recipient type
				
				$patwrong = $false
				$YN = $null
				do
				{
					do
					{
						#show options for recipient type detail selection
						$RecType = Read-Host "Please select the recipient type(s) you want to include. Separate multiple values by comma (1,2,...).

        [1] User Mailbox 
        [2] Shared Mailbox
        [3] Room Mailbox
        [4] Team Mailbox
        [5] Group Mailbox
        [C] Cancel

        Your selection"
						""
						
						if ($RecType -eq 'c')
						{
							"Exiting..."
							return
						}
						#Verify entered value
						$pattern = '^(?!.*?([1-5]).*?\1)[1-5](?:,[1-5])*$'
						
						if ($RecType.Length -gt 9 -or $RecType -notmatch $pattern)
						{
							'Incorrect format!'
							sleep -Seconds 1
							$patwrong = $true
							
							#return
						}
						else
						{
							$patwrong = $false
						}
					}
					until ($patwrong -eq $false)
					#Create string for get-mailbox -recipienttypedetails parameter according to user selection
					$RecType = $RecType.Replace('1', 'UserMailbox').Replace('2', 'SharedMailbox').Replace('3', 'RoomMailbox').Replace('4', 'TeamMailbox').Replace('5', 'GroupMailbox')
					#Ask if selected types are correct
					"Following recipient type(s) will be included:"
					""
					$($RecType -split ',')
					""
					
					$YN = read-host "Correct? [Y/N]"
				}
				until ($yn -eq 'y')
				
				#endregion
				
				#region set extension length
				cls
				#Ask for length of extension number
				$ExtLen = Read-Host "Please enter the length of the extension number in your environment for UM"
				#Check if a valid digit was entered
				if ($ExtLen -notmatch "\d" -or $ExtLen -eq 0)
				{
					Write-Host "Unsupported format. Only digits greater then 0 are supported." -ForegroundColor Red
					return
				}
				#endregion
				
				#region O365 | Connect
				if ($ExcOpt -eq 1)
				{
					#Connect to Exchange Online remotely
					try
					{
						$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop
						Write-Host "Connecting to Exchange Online..." -ForegroundColor Green
						Import-PSSession $Session -ErrorAction Stop | Out-Null
						(Get-Date -Format G) + " " + "Exchange Online connected" | Tee-Object -FilePath $log -Append
					}
					catch
					{
						(Get-Date -Format G) + " " + "ERROR: " + $_.exception.message | Tee-Object -FilePath $log -Append
						return
					}
					#Connect to AzureAD
					try
					{
						Write-Host "Connecting to Azure AD..." -ForegroundColor Green
						
						if ($AADModule -eq 'AAD')
						{
							#Use AzureAD module
							$aad = Connect-AzureAD -Credential $O365Creds -ErrorAction Stop
							
						}
						else
						{
							#Use MSonline module
							connect-MsolService -credential $O365Creds -ErrorAction Stop
							
						}
						
						(Get-Date -Format G) + " " + "Azure AD $($aad.TenantDomain) connected" | Tee-Object -FilePath $log -Append
						
					}
					catch
					{
						(Get-Date -Format G) + " " + "ERROR: " + $_.exception.message | Tee-Object -FilePath $log -Append
						return
					}
					
				}
				#endregion
				
				#region On-Premises | connect to Exchange
				if ($ExcOpt -eq 2)
				{
					#Ask for Exchange server name
					$Exchange = Read-Host "Enter FQDN, or short name of on-premises Exchange server. E.g. ""EXCSRV01.contoso.com, or EXCSRV01"
					
					try
					{
						#Remote connect to Exchange On-Prem 
						$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchange/PowerShell/ -Authentication Kerberos -Credential $OnPremCreds
						Import-PSSession $Session -ErrorAction Stop | Out-Null
						(Get-Date -Format G) + " " + "Exchange connected" | Tee-Object -FilePath $log -Append
					}
					catch
					{
						(Get-Date -Format G) + " " + "ERROR: " + $_.exception.message | Tee-Object -FilePath $log -Append
						Return
					}
					
				}
				#endregion
				
				#region Select UMPolicy
				cls
				#Get all UM policies
				$UMPolicies = Get-UMMailboxPolicy
				
				#Count and list UM policies 
				$count = 0
				foreach ($policy in $UMPolicies)
				{
					$count++
					Write-Host "[$count] - $($policy.Name)" -ForegroundColor Cyan
					
				}
				#Ask for UM policy to choose. (Enter number)
				[INT]$Idx = Read-Host "Enter number of UM policy to choose"
				
				#Check if entered number is valid
				if ($Idx -eq 0 -or $Idx -gt $count)
				{
					#If entered number is not valid, end script
					cls
					Write-Host "Please select a number between 1 and $count. Script ends now." -ForegroundColor Red
					return
				}
				else
				{
					#Select UM policy on base of entered number
					$UMPolicy = $UMPolicies[$idx - 1].Name
					cls
					Write-Host "You have selected the following policy: $UMPolicy" -BackgroundColor Blue
				}
				#endregion
				
				#region Enable UM users
				Write-Host "Fetching mailboxes..." -ForegroundColor Green
				#Get all mailboxes where UM is not enabled
				$mbxs = get-mailbox -RecipientTypeDetails $RecType -ResultSize unlimited | where { $_.UMEnabled -eq $false }
				$mcount = 0
				$successcount = 0
				$errorcount = 0
				
				#Go through all found mail boxes
				foreach ($mbx in $mbxs)
				{
					#Create progress bar
					$mcount++
					$percent = "{0:N1}" -f ($mcount / $mbxs.count * 100)
					Write-Progress -Activity "Enabling UM" -status "Enabling Service for $($mbx.PrimarySMTPAddress)" -percentComplete $percent -CurrentOperation "Percent completed: $percent% (no. $mcount) of $($mbxs.count) mailboxes"
					
					#Get phone number of user
					try
					{
						switch ($ExcOpt)
						{
							1 {
								#If O365 selected, get phone number from Azure AD 
								if ($AADModule -eq 'AAD')
								{
									#Use AzureAD module
									$aadUser = get-azureADUser -SearchString $mbx.UserPrincipalName -erroraction Stop
									$phone = $aadUser.TelephoneNumber
									#Throw an error if no phone number was found for the user
									if ($phone -eq "" -or $phone -eq $null)
									{
										throw "$($mbx.UserPrincipalName) - No phone number found"
										return
									}
									if ($aadUser.AssignedPlans.service -notcontains 'MicrosoftCommunicationsOnline')
									{
										throw "Error: $($mbx.userprincipalname) has no S4B Online (Plan 2) plan assigned."
										return
									}
									if ($aadUser.AssignedPlans.service -notcontains 'exchange')
									{
										throw "Error: $($mbx.userprincipalname) has no Exchange Online (E1, or E2) plan assigned."
										return
									}
								}
								else
								{
									#Use MSOnline module
									$aadUser = get-MsolUser -SearchString $mbx.UserPrincipalName -erroraction Stop
									$phone = $aadUser.PhoneNumber
									#Throw an error if no phone number was found for the user
									if ($phone -eq "" -or $phone -eq $null)
									{
										throw "$($mbx.UserPrincipalName) - No phone number found"
										return
									}
								}
							}
							2 {
								#If On-Prem is selected, use ADSI searcher to get the phone number
								$b = [adsisearcher]::new("userprincipalname=$($mbx.UserPrincipalName)")
								$result = $b.FindOne()
								$phone = $result.Properties.telephonenumber
								#Throw an error if no phone number was found for the user
								if ($phone -eq "" -or $phone -eq $null)
								{
									throw "$($mbx.UserPrincipalName) - No phone number found"
									return
								}
								
							}
						}
						
						
						#LineURI string modifiy for extension number (get only the last digits that were defined in the beginning)
						$str = $phone.TrimStart("tel:+").replace(" ", "") #Trim all spaces
						$length = $str.Length #Get length of the string
						$URI = $str.Substring(($length - $ExtLen)) #Select only substring starting from string length minus defined length
						
						#Create extension mapping (maybe used for future versions)
						$ExtensionMap = @{
							User = $mbx.PrimarySMTPAddress
							Extension = $URI
						}
						#Enable UM for the mailbox
						Enable-UMMailbox -Identity $mbx.PrimarySMTPAddress -UMMailboxPolicy $UMPolicy -SIPResourceIdentifier $mbx.PrimarySMTPAddress`
										 -Extensions $ExtensionMap.Extension -PinExpired $false -ErrorAction Stop #-WhatIf
						
						#Log
						$datetime = (Get-Date -Format G)
						"$datetime SUCCESS: $($mbx.UserPrincipalName) has been enabled for UM" | Tee-Object $log -Append
						
						#Count successfully enabled mailboxes
						$successcount++
						
					}
					catch
					{
						#Log error
						$datetime = (Get-Date -Format G)
						"$datetime ERROR: $($mbx.UserPrincipalName)  $($_.Exception.Message)" | Tee-Object $log -Append
						#Count errors
						$errorcount++
					}
					
					
				}
				#End progress bar
				Write-Progress -Activity "Enabling UM" -Completed
				
				#endregion
				
				#region show summary
				#Number of successes
				Write-Host "$successcount of $($mbxs.count) mailboxes have been successfully enabled for UM! " -ForegroundColor Green
				#If errors occurred show number of errors
				if ($errorcount -gt 0)
				{
					Write-Host "Number of errors during execution: $errorcount. Please check the log ""$log"" for details." -ForegroundColor Green
				}
				"Press any key to exit"
				cmd /c pause | Out-Null
				#endregion
			}
			#endregion   
			#region Option 52) Remove  old EAS devices
			52 {
				{
					$age = $null
					
					#user input age
					$pattern = "\d+"
					Do
					{
						$Age = Read-Host "Please specify max number of days. Older entries will be removed (leave empty to cancel)"
						cls
						if ($age -notmatch $pattern -and $age -ne "")
						{
							write-host "Please enter a valid number in number format, or ""C"" to cancel!" -ForegroundColor Yellow
							sleep -Seconds 2
							cls
						}
						
					}
					Until ($age -eq "" -or $age -match $pattern -or $age -eq "c")
					if ($age -eq "")
					{
						"Cancelled"
						return
					}
					
					
					
					# Variables
					$now = Get-Date #Used for timestamps
					$date = $now.ToShortDateString() #Short date format for email message subject
					
					$report = @()
					
					$stats = @("DeviceID",
						"DeviceAccessState",
						"DeviceAccessStateReason",
						"DeviceModel"
						"DeviceType",
						"DeviceFriendlyName",
						"DeviceOS",
						"LastSyncAttemptTime",
						"LastSuccessSync"
					)
					
					$reportemailsubject = "Exchange ActiveSync Device Report - $date"
					$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
					$reportfile = "$myDir\ExchangeActiveSyncDevice-ToDelete.csv"
					
					
					
					
					
					# Initialize
					#Add Exchange 2010/2013/2016 snapin if not already loaded in the PowerShell session
					if (!(Get-PSSnapin | where { $_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010" }))
					{
						try
						{
							Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
						}
						catch
						{
							#Snapin was not loaded
							Write-Warning $_.Exception.Message
							EXIT
						}
						. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
						Connect-ExchangeServer -auto -AllowClobber
					}
					
					
					
					# Script
					Write-Host "keep it simple but significant" -ForegroundColor magenta
					Start-Sleep -s 2
					Write-Host "Fetching List of Mailboxes with EAS Device partnerships" -ForegroundColor cyan
					Start-Sleep -s 5
					Write-Host "Don't worry, this can take a while..." -ForegroundColor cyan
					
					$MailboxesWithEASDevices = @(Get-CASMailbox -Resultsize Unlimited | Where { $_.HasActiveSyncDevicePartnership })
					
					Write-Host "$($MailboxesWithEASDevices.count) mailboxes with EAS device partnerships"
					
					$i = 0
					
					Foreach ($Mailbox in $MailboxesWithEASDevices)
					{
						
						$EASDeviceStats = @(Get-ActiveSyncDeviceStatistics -Mailbox $Mailbox.Identity -WarningAction SilentlyContinue)
						
						Write-Host "$($Mailbox.Identity) has $($EASDeviceStats.Count) device(s)"
						
						$MailboxInfo = Get-Mailbox $Mailbox.Identity | Select DisplayName, PrimarySMTPAddress, OrganizationalUnit
						
						Foreach ($EASDevice in $EASDeviceStats)
						{
							Write-Host -ForegroundColor Green "Processing $($EASDevice.DeviceID)"
							
							$lastsyncattempt = ($EASDevice.LastSyncAttemptTime)
							
							if ($lastsyncattempt -eq $null)
							{
								$syncAge = "Never"
							}
							else
							{
								$syncAge = ($now - $lastsyncattempt).Days
							}
							
							#Add to report if last sync attempt greater than Age specified
							if ($syncAge -ge $Age -or $syncAge -eq "Never" -and $EASDevice.DeviceID -ne 0)
							{
								Write-Host -ForegroundColor Yellow "$($EASDevice.DeviceID) sync age of $syncAge days is greater than $age, adding to report"
								
								$reportObj = New-Object PSObject
								$reportObj | Add-Member NoteProperty -Name "Display Name" -Value $MailboxInfo.DisplayName
								$reportObj | Add-Member NoteProperty -Name "Organizational Unit" -Value $MailboxInfo.OrganizationalUnit
								$reportObj | Add-Member NoteProperty -Name "Email Address" -Value $MailboxInfo.PrimarySMTPAddress
								$reportObj | Add-Member NoteProperty -Name "Sync Age (Days)" -Value $syncAge
								$reportObj | Add-Member NoteProperty -Name "GUID" -Value $EASDevice.GUID
								
								Foreach ($stat in $stats)
								{
									$reportObj | Add-Member NoteProperty -Name $stat -Value $EASDevice.$stat
								}
								
								$report += $reportObj
							}
						}
						$i++
						Write-Progress -activity "Gethering EAS devices . . ." -status "Collected: $i of $($MailboxesWithEASDevices.Count)" -percentComplete (($i / $MailboxesWithEASDevices.Count) * 100)
					}
					Write-Progress -activity "Gethering EAS devices . . ." -Completed
					
					Write-Host -ForegroundColor White "Saving report to $reportfile"
					$report | Export-Csv -NoTypeInformation $reportfile -Encoding UTF8
					
					ii $reportfile #Open the CSV. File 
					Write-Host "!!! with great power comes great responsibility !!!" -ForegroundColor magenta
					Write-Host "Check the CSV File before you continue! To continue push ENTER" -ForegroundColor Gray -NoNewline
					$dummy = Read-Host
					
					$ReportToDelete = Import-csv $reportfile
					###
					
					$counter = 0
					$sum = $ReportToDelete.count
					foreach ($i in $ReportToDelete)
					{
						try
						{
							write-host $i."Display Name" $i."LastSuccessSync" $i."DeviceFriendlyName"
							Remove-MobileDevice -Identity $i."GUID" -Confirm:$false -erroraction Stop #Remove the selected MobileDevices (by GUID)
							Write-Host "Device removed" -ForegroundColor Green
							(get-date -Format g) + " Success: Removed device: " + $i."Display Name" + $i."DeviceFriendlyName" | Out-File $PSScriptRoot\Successlog.log -Append
						}
						catch
						{
							
							(get-date -Format g) + " Error: " + $i."Display Name" + $i."DeviceFriendlyName" + " " + $_.exception.message | Out-File $PSScriptRoot\errorlog.log -Append
							Write-Host "Error while removing device" -ForegroundColor Red
							$_.exception.message
						}
						$counter++
						Write-Progress -activity "Removing EAS devices . . ." -status "Processed: $counter of $($sum)" -percentComplete (($counter / $sum) * 100)
					}
					Write-Progress -Activity "Removing EAS devices . . ." -Completed
					
					Write-Host "Active sync Devices older then $Age Days are successfully removed for the Exchange Organization" -ForegroundColor green
				}
			}
			#endregion   
			#region Option 53) Deploy Microsoft Teams Desktop Client
			53 {
				"This function is not yet implemented"
				#      Deploy Microsoft Teams Desktop Client
                 <#          <#
.SYNOPSIS
Install-MicrosoftTeams.ps1 - Microsoft Teams Desktop Client Deployment Script

.DESCRIPTION 
This PowerShell script will silently install the Microsoft Teams desktop client.

The Teams client installer can be downloaded from Microsoft:
https://teams.microsoft.com/downloads

.PARAMETER SourcePath
Specifies the source path for the Microsoft Teams installer.


.EXAMPLE
.\Install-MicrosoftTeams.ps1 -Source \\mgmt\Installs\MicrosoftTeams

Installs the Microsoft Teams client from the Installs share on the server MGMT.

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:   http://paulcunningham.me
* Twitter:   https://twitter.com/paulcunningham
* LinkedIn:  http://au.linkedin.com/in/cunninghamp/
* Github:    https://github.com/cunninghamp

For more Office 365 tips, tricks and news
check out Practical 365.

* Website:   http://practical365.com
* Twitter:   http://twitter.com/practical365

Change Log
V1.00, 15/03/2017 - Initial version
#>
				
				#requires -version 4
				{
					
					[CmdletBinding()]
					param (
						
						[Parameter(Mandatory = $true)]
						[string]$SourcePath
						
					)
					
					
					function DoInstall
					{
						
						$Installer = "$($SourcePath)\Teams_windows_x64.exe"
						
						If (!(Test-Path $Installer))
						{
							throw "Unable to locate Microsoft Teams client installer at $($installer)"
						}
						
						Write-Host "Attempting to install Microsoft Teams client"
						
						try
						{
							$process = Start-Process -FilePath "$Installer" -ArgumentList "-s" -Wait -PassThru -ErrorAction STOP
							
							if ($process.ExitCode -eq 0)
							{
								Write-Host -ForegroundColor Green "Microsoft Teams setup started without error."
							}
							else
							{
								Write-Warning "Installer exit code  $($process.ExitCode)."
							}
						}
						catch
						{
							Write-Warning $_.Exception.Message
						}
						
					}
					
					#Check if Office is already installed, as indicated by presence of registry key
					$installpath = "$($env:LOCALAPPDATA)\Microsoft\Teams"
					
					if (-not (Test-Path "$($installpath)\Update.exe"))
					{
						DoInstall
					}
					else
					{
						if (Test-Path "$($installpath)\.dead")
						{
							Write-Host "Teams was previously installed but has been uninstalled. Will reinstall."
							DoInstall
						}
					}
					
					
					
					
					
					
					$Reboot = $true
				} #>
			}
			#endregion   
			#region Option 54) Order certificate >>GO DADDY<<
			54 {
				"This function is not yet implemented"
				<#
				#      Order certificate >>GO DADDY<<
				#Order certificate
				empty
				
				#Variables
				$cersrv = Read-Host 'Enter the server name where you wanna import the certificate <e.g. EXCsrv01>'
				$cerpath = Read-Host 'Enter the the path, where your certificate is located <e.g. \\FileServer01\Data\>'
				$cercertname = Read-Host 'Enter the Name of the certificate <e.g. 'Exported Fabrikam Cert.pfx'>'
				$cerPW = Read-Host 'Enter the Password of the .PFX file PLEASE NOTE, YOU ENTER IT AS PLAIN TEXT! the password will be converted to a Secure String automaticaly!'
				
				#Script import certificate
				Import-ExchangeCertificate -Server $cersrv -FileName "$cerpath", "$cercertname" -Password (ConvertTo-SecureString -String $cerPW -AsPlainText -Force)
				
				$Reboot = $false#>
			}
			#endregion
			#region Option 55) Order certificate >>DIGICERT<<
			55 {
				"This function is not yet implemented"
				#      Order certificate >>DIGICERT<<
				#Order certificate
				<#empty
				
				#Variables
				$cersrv = Read-Host 'Enter the server name where you wanna import the certificate <e.g. EXCsrv01>'
				$cerpath = Read-Host 'Enter the the path, where your certificate is located <e.g. \\FileServer01\Data\>'
				$cercertname = Read-Host 'Enter the Name of the certificate <e.g. 'Exported Fabrikam Cert.pfx'>'
				$cerPW = Read-Host 'Enter the Password of the .PFX file PLEASE NOTE, YOU ENTER IT AS PLAIN TEXT! the password will be converted to a Secure String automaticaly!'
				
				#Script import certificate
				Import-ExchangeCertificate -Server $cersrv -FileName "$cerpath", "$cercertname" -Password (ConvertTo-SecureString -String $cerPW -AsPlainText -Force)
				
				$Reboot = $false#>
			}
			#endregion
			#region Option 60) Generate Health Report for an Exchange Server 2016/2013/2010 Environment
			60 {
				generateHealthReport
			}
			#endregion   
			#region Option 61) Generate Exchange Environment Reports
			61 {
				#      Generate Exchange Environment Reports
                           <#
    .SYNOPSIS
    Creates a HTML Report describing the Exchange environment 
   
       Steve Goodman
       
       THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
       RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
       
       Version 1.6.2 January 2017
       
    .DESCRIPTION
       
    This script creates a HTML report showing the following information about an Exchange 
    2016, 2013, 2010 and to a lesser extent, 2007 and 2003, environment. 
    
    The following is shown:
       
       * Report Generation Time
       * Total Servers per Exchange Version (2003 > 2010 or 2007 > 2016)
       * Total Mailboxes per Exchange Version, Office 365 and Organisation
       * Total Roles in the environment
             
       Then, per site:
       * Total Mailboxes per site
    * Internal, External and CAS Array Hostnames
       * Exchange Servers with:
             o Exchange Server Version
             o Service Pack
             o Update Rollup and rollup version
             o Roles installed on server and mailbox counts
             o OS Version and Service Pack
             
       Then, per Database availability group (Exchange 2010/2013/2016):
       * Total members per DAG
       * Member list
       * Databases, detailing:
             o Mailbox Count and Average Size
             o Archive Mailbox Count and Average Size (Only shown if DAG includes Archive Mailboxes)
             o Database Size and whitespace
             o Database and log disk free
             o Last Full Backup (Only shown if one or more DAG database has been backed up)
             o Circular Logging Enabled (Only shown if one or more DAG database has Circular Logging enabled)
             o Mailbox server hosting active copy
             o List of mailbox servers hosting copies and number of copies
             
       Finally, per Database (Non DAG DBs/Exchange 2007/Exchange 2003)
       * Databases, detailing:
             o Storage Group (if applicable) and DB name
             o Server hosting database
             o Mailbox Count and Average Size
             o Archive Mailbox Count and Average Size (Only shown if DAG includes Archive Mailboxes)
             o Database Size and whitespace
             o Database and log disk free
             o Last Full Backup (Only shown if one or more DAG database has been backed up)
             o Circular Logging Enabled (Only shown if one or more DAG database has Circular Logging enabled)
             
       This does not detail public folder infrastructure, or examine Exchange 2007/2003 CCR/SCC clusters
       (although it attempts to detect Clustered Exchange 2007/2003 servers, signified by ClusMBX).
       
       IMPORTANT NOTE: The script requires WMI and Remote Registry access to Exchange servers from the server 
       it is run from to determine OS version, Update Rollup, Exchange 2007/2003 cluster and DB size information.
       
       .PARAMETER HTMLReport
    Filename to write HTML Report to
       
       .PARAMETER SendMail
       Send Mail after completion. Set to $True to enable. If enabled, -MailFrom, -MailTo, -MailServer are mandatory
       
       .PARAMETER MailFrom
       Email address to send from. Passed directly to Send-MailMessage as -From
       
       .PARAMETER MailTo
       Email address to send to. Passed directly to Send-MailMessage as -To
       
       .PARAMETER MailServer
       SMTP Mail server to attempt to send through. Passed directly to Send-MailMessage as -SmtpServer
       
       .PARAMETER ScheduleAs
       Attempt to schedule the command just executed for 10PM nightly. Specify the username here, schtasks (under the hood) will ask for a password later.
    
       .PARAMETER ViewEntireForest
       By default, true. Set the option in Exchange 2007 or 2010 to view all Exchange servers and recipients in the forest.
   
    .PARAMETER ServerFilter
       Use a text based string to filter Exchange Servers by, e.g. NL-* -  Note the use of the wildcard (*) character to allow for multiple matches.
    
       .EXAMPLE
    Generate the HTML report 
    .\Get-ExchangeEnvironmentReport.ps1 -HTMLReport .\report.html
       
    #>
				#region user input
				$HTMLReport = Read-Host "Specifiy alternate path and name for report file. Default is ""exchangeenvironmentreport.html"""
				if ($HTMLReport = "")
				{
					$ReportFile = "exchangeenvironmentreport.html"
				}
				$SendMailYesNo = Read-Host "Send e-mail with report? [Y/N] Default is [N]"
				
				switch ($SendMailYesNo)
				{
					Y{ $SendEmail = $true }
					N{ $SendEmail = $false }
					default { "No option selected. Exiting"; Return }
				}
				if ($SendEmail)
				{
					$AlertsOnlyYN = Read-Host "Send email only if error or warning was detected?[Y/N] Default is [N]"
					switch ($AlertsOnlyYN)
					{
						Y{ $AlertsOnly = $true }
						N{ $AlertsOnly = $false }
						default { $AlertsOnly = $false }
					}
					$MailServer = Read-Host "Enter SMTP Server"
					$MailTo = Read-Host -Prompt "Enter recipients SMTP address"
					$MailFrom = Read-Host -Prompt "Enter senders SMTP address"
					
				}
				$ScheduleAsYN = Read-Host "Do you want to schedule the execution?[Y/N]"
				if ($ScheduleAsYN = "Y")
				{
					$ScheduleAs = Read-Host "Enter username in which context the scheduled will run"
					if ($ScheduleAs = "")
					{
						Write-Host "No username specified. Exiting now."
						Return
					}
				}
				$ViewEntireForestYN = Read-Host "View entire forest (all Exchange servers and recipients)[Y/N] Default is [Y]"
				switch ($ViewEntireForestYN)
				{
					Y{ $ViewEntireForest = $true }
					N{ $ViewEntireForest = $false }
					default { $AlertsOnly = $true }
				}
				$ServerFilter = Read-Host "Specifiy a filter for server names. Wildcards are allowed. Default is ""*"""
				if ($ServerFilter = "")
				{
					$ServerFilter = "*"
				}
				#endregion
				
				# Sub-Function to Get Database Information. Shorter than expected..
				function _GetDAG
				{
					param ($DAG)
					@{
						Name = $DAG.Name.ToUpper()
						MemberCount = $DAG.Servers.Count
						Members = [array]($DAG.Servers | % { $_.Name })
						Databases = @()
					}
				}
				
				
				# Sub-Function to Get Database Information
				function _GetDB
				{
					param ($Database,
						$ExchangeEnvironment,
						$Mailboxes,
						$ArchiveMailboxes,
						$E2010)
					
					# Circular Logging, Last Full Backup
					if ($Database.CircularLoggingEnabled) { $CircularLoggingEnabled = "Yes" }
					else { $CircularLoggingEnabled = "No" }
					if ($Database.LastFullBackup) { $LastFullBackup = $Database.LastFullBackup.ToString() }
					else { $LastFullBackup = "Not Available" }
					
					# Mailbox Average Sizes
					$MailboxStatistics = [array]($ExchangeEnvironment.Servers[$Database.Server.Name].MailboxStatistics | Where { $_.Database -eq $Database.Identity })
					if ($MailboxStatistics)
					{
						[long]$MailboxItemSizeB = 0
						$MailboxStatistics | %{ $MailboxItemSizeB += $_.TotalItemSizeB }
						[long]$MailboxAverageSize = $MailboxItemSizeB / $MailboxStatistics.Count
					}
					else
					{
						$MailboxAverageSize = 0
					}
					
					# Free Disk Space Percentage
					if ($ExchangeEnvironment.Servers[$Database.Server.Name].Disks)
					{
						foreach ($Disk in $ExchangeEnvironment.Servers[$Database.Server.Name].Disks)
						{
							if ($Database.EdbFilePath.PathName -like "$($Disk.Name)*")
							{
								$FreeDatabaseDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
							}
							if ($Database.ExchangeVersion.ExchangeBuild.Major -ge 14)
							{
								if ($Database.LogFolderPath.PathName -like "$($Disk.Name)*")
								{
									$FreeLogDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
								}
							}
							else
							{
								$StorageGroupDN = $Database.DistinguishedName.Replace("CN=$($Database.Name),", "")
								$Adsi = [adsi]"LDAP://$($Database.OriginatingServer)/$($StorageGroupDN)"
								if ($Adsi.msExchESEParamLogFilePath -like "$($Disk.Name)*")
								{
									$FreeLogDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
								}
							}
						}
					}
					else
					{
						$FreeLogDiskSpace = $null
						$FreeDatabaseDiskSpace = $null
					}
					
					if ($Database.ExchangeVersion.ExchangeBuild.Major -ge 14 -and $E2010)
					{
						# Exchange 2010 Database Only
						$CopyCount = [int]$Database.Servers.Count
						if ($Database.MasterServerOrAvailabilityGroup.Name -ne $Database.Server.Name)
						{
							$Copies = [array]($Database.Servers | % { $_.Name })
						}
						else
						{
							$Copies = @()
						}
						# Archive Info
						$ArchiveMailboxCount = [int]([array]($ArchiveMailboxes | Where { $_.ArchiveDatabase -eq $Database.Name })).Count
						$ArchiveStatistics = [array]($ArchiveMailboxes | Where { $_.ArchiveDatabase -eq $Database.Name } | Get-MailboxStatistics -Archive)
						if ($ArchiveStatistics)
						{
							[long]$ArchiveItemSizeB = 0
							$ArchiveStatistics | %{ $ArchiveItemSizeB += $_.TotalItemSize.Value.ToBytes() }
							[long]$ArchiveAverageSize = $ArchiveItemSizeB / $ArchiveStatistics.Count
						}
						else
						{
							$ArchiveAverageSize = 0
						}
						# DB Size / Whitespace Info
						[long]$Size = $Database.DatabaseSize.ToBytes()
						[long]$Whitespace = $Database.AvailableNewMailboxSpace.ToBytes()
						$StorageGroup = $null
						
					}
					else
					{
						$ArchiveMailboxCount = 0
						$CopyCount = 0
						$Copies = @()
						# 2003 & 2007, Use WMI (Based on code by Gary Siepser, http://bit.ly/kWWMb3)
						$Size = [long](get-wmiobject cim_datafile -computername $Database.Server.Name -filter ('name=''' + $Database.edbfilepath.pathname.replace("\", "\\") + '''')).filesize
						if (!$Size)
						{
							Write-Warning "Cannot detect database size via WMI for $($Database.Server.Name)"
							[long]$Size = 0
							[long]$Whitespace = 0
						}
						else
						{
							[long]$MailboxDeletedItemSizeB = 0
							if ($MailboxStatistics)
							{
								$MailboxStatistics | %{ $MailboxDeletedItemSizeB += $_.TotalDeletedItemSizeB }
							}
							$Whitespace = $Size - $MailboxItemSizeB - $MailboxDeletedItemSizeB
							if ($Whitespace -lt 0) { $Whitespace = 0 }
						}
						$StorageGroup = $Database.DistinguishedName.Split(",")[1].Replace("CN=", "")
					}
					
					@{
						Name = $Database.Name
						StorageGroup = $StorageGroup
						ActiveOwner = $Database.Server.Name.ToUpper()
						MailboxCount = [long]([array]($Mailboxes | Where { $_.Database -eq $Database.Identity })).Count
						MailboxAverageSize = $MailboxAverageSize
						ArchiveMailboxCount = $ArchiveMailboxCount
						ArchiveAverageSize = $ArchiveAverageSize
						CircularLoggingEnabled = $CircularLoggingEnabled
						LastFullBackup = $LastFullBackup
						Size = $Size
						Whitespace = $Whitespace
						Copies = $Copies
						CopyCount = $CopyCount
						FreeLogDiskSpace = $FreeLogDiskSpace
						FreeDatabaseDiskSpace = $FreeDatabaseDiskSpace
					}
				}
				
				
				# Sub-Function to get mailbox count per server.
				# New in 1.5.2
				function _GetExSvrMailboxCount
				{
					param ($Mailboxes,
						$ExchangeServer,
						$Databases)
					# The following *should* work, but it doesn't. Apparently, ServerName is not always returned correctly which may be the cause of
					# reports of counts being incorrect
					#([array]($Mailboxes | Where {$_.ServerName -eq $ExchangeServer.Name})).Count
					
					# ..So as a workaround, I'm going to check what databases are assigned to each server and then get the mailbox counts on a per-
					# database basis and return the resulting total. As we already have this information resident in memory it should be cheap, just
					# not as quick.
					$MailboxCount = 0
					foreach ($Database in [array]($Databases | Where { $_.Server -eq $ExchangeServer.Name }))
					{
						$MailboxCount += ([array]($Mailboxes | Where { $_.Database -eq $Database.Identity })).Count
					}
					$MailboxCount
					
				}
				
				# Sub-Function to Get Exchange Server information
				function _GetExSvr
				{
					param ($E2010,
						$ExchangeServer,
						$Mailboxes,
						$Databases,
						$Hybrids)
					
					# Set Basic Variables
					$MailboxCount = 0
					$RollupLevel = 0
					$RollupVersion = ""
					$ExtNames = @()
					$IntNames = @()
					$CASArrayName = ""
					
					# Get WMI Information
					$tWMI = Get-WmiObject Win32_OperatingSystem -ComputerName $ExchangeServer.Name -ErrorAction SilentlyContinue
					if ($tWMI)
					{
						$OSVersion = $tWMI.Caption.Replace("(R)", "").Replace("Microsoft ", "").Replace("Enterprise", "Ent").Replace("Standard", "Std").Replace(" Edition", "")
						$OSServicePack = $tWMI.CSDVersion
						$RealName = $tWMI.CSName.ToUpper()
					}
					else
					{
						Write-Warning "Cannot detect OS information via WMI for $($ExchangeServer.Name)"
						$OSVersion = "N/A"
						$OSServicePack = "N/A"
						$RealName = $ExchangeServer.Name.ToUpper()
					}
					$tWMI = Get-WmiObject -query "Select * from Win32_Volume" -ComputerName $ExchangeServer.Name -ErrorAction SilentlyContinue
					if ($tWMI)
					{
						$Disks = $tWMI | Select Name, Capacity, FreeSpace | Sort-Object -Property Name
					}
					else
					{
						Write-Warning "Cannot detect OS information via WMI for $($ExchangeServer.Name)"
						$Disks = $null
					}
					
					# Get Exchange Version
					if ($ExchangeServer.AdminDisplayVersion.Major -eq 6)
					{
						$ExchangeMajorVersion = "$($ExchangeServer.AdminDisplayVersion.Major).$($ExchangeServer.AdminDisplayVersion.Minor)"
						$ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.FilePatchLevelDescription.Replace("Service Pack ", "")
					}
					elseif ($ExchangeServer.AdminDisplayVersion.Major -eq 15 -and $ExchangeServer.AdminDisplayVersion.Minor -eq 1)
					{
						$ExchangeMajorVersion = [double]"$($ExchangeServer.AdminDisplayVersion.Major).$($ExchangeServer.AdminDisplayVersion.Minor)"
						$ExchangeSPLevel = 0
					}
					else
					{
						$ExchangeMajorVersion = $ExchangeServer.AdminDisplayVersion.Major
						$ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.Minor
					}
					# Exchange 2007+
					if ($ExchangeMajorVersion -ge 8)
					{
						# Get Roles
						$MailboxStatistics = $null
						[array]$Roles = $ExchangeServer.ServerRole.ToString().Replace(" ", "").Split(",");
						# Add Hybrid "Role" for report
						if ($Hybrids -contains $ExchangeServer.Name)
						{
							$Roles += "Hybrid"
						}
						if ($Roles -contains "Mailbox")
						{
							$MailboxCount = _GetExSvrMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
							if ($ExchangeServer.Name.ToUpper() -ne $RealName)
							{
								$Roles = [array]($Roles | Where { $_ -ne "Mailbox" })
								$Roles += "ClusteredMailbox"
							}
							# Get Mailbox Statistics the normal way, return in a consitent format
							$MailboxStatistics = Get-MailboxStatistics -Server $ExchangeServer | Select DisplayName, @{ Name = "TotalItemSizeB"; Expression = { $_.TotalItemSize.Value.ToBytes() } }, @{ Name = "TotalDeletedItemSizeB"; Expression = { $_.TotalDeletedItemSize.Value.ToBytes() } }, Database
						}
						# Get HTTPS Names (Exchange 2010 only due to time taken to retrieve data)
						if ($Roles -contains "ClientAccess" -and $E2010)
						{
							
							Get-OWAVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames += $_.ExternalURL.Host; $IntNames += $_.InternalURL.Host; }
							Get-WebServicesVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames += $_.ExternalURL.Host; $IntNames += $_.InternalURL.Host; }
							Get-OABVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames += $_.ExternalURL.Host; $IntNames += $_.InternalURL.Host; }
							Get-ActiveSyncVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames += $_.ExternalURL.Host; $IntNames += $_.InternalURL.Host; }
							if (Get-Command Get-MAPIVirtualDirectory -ErrorAction SilentlyContinue)
							{
								Get-MAPIVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames += $_.ExternalURL.Host; $IntNames += $_.InternalURL.Host; }
							}
							if (Get-Command Get-ClientAccessService -ErrorAction SilentlyContinue)
							{
								$IntNames += (Get-ClientAccessService -Identity $ExchangeServer.Name).AutoDiscoverServiceInternalURI.Host
							}
							else
							{
								$IntNames += (Get-ClientAccessServer -Identity $ExchangeServer.Name).AutoDiscoverServiceInternalURI.Host
							}
							
							if ($ExchangeMajorVersion -ge 14)
							{
								Get-ECPVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames += $_.ExternalURL.Host; $IntNames += $_.InternalURL.Host; }
							}
							$IntNames = $IntNames | Sort-Object -Unique
							$ExtNames = $ExtNames | Sort-Object -Unique
							$CASArray = Get-ClientAccessArray -Site $ExchangeServer.Site.Name
							if ($CASArray)
							{
								$CASArrayName = $CASArray.Fqdn
							}
						}
						
						# Rollup Level / Versions (Thanks to Bhargav Shukla http://bit.ly/msxGIJ)
						if ($ExchangeMajorVersion -ge 14)
						{
							$RegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches"
						}
						else
						{
							$RegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\461C2B4266EDEF444B864AD6D9E5B613\\Patches"
						}
						$RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name);
						if ($RemoteRegistry)
						{
							$RUKeys = $RemoteRegistry.OpenSubKey($RegKey).GetSubKeyNames() | ForEach { "$RegKey\\$_" }
							if ($RUKeys)
							{
								[array]($RUKeys | %{ $RemoteRegistry.OpenSubKey($_).getvalue("DisplayName") }) | %{
									if ($_ -like "Update Rollup *")
									{
										$tRU = $_.Split(" ")[2]
										if ($tRU -like "*-*") { $tRUV = $tRU.Split("-")[1]; $tRU = $tRU.Split("-")[0] }
										else { $tRUV = "" }
										if ([int]$tRU -ge [int]$RollupLevel) { $RollupLevel = $tRU; $RollupVersion = $tRUV }
									}
								}
							}
						}
						else
						{
							Write-Warning "Cannot detect Rollup Version via Remote Registry for $($ExchangeServer.Name)"
						}
						# Exchange 2013 CU or SP Level
						if ($ExchangeMajorVersion -ge 15)
						{
							$RegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Microsoft Exchange v15"
							$RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name);
							if ($RemoteRegistry)
							{
								$ExchangeSPLevel = $RemoteRegistry.OpenSubKey($RegKey).getvalue("DisplayName")
								if ($ExchangeSPLevel -like "*Service Pack*" -or $ExchangeSPLevel -like "*Cumulative Update*")
								{
									$ExchangeSPLevel = $ExchangeSPLevel.Replace("Microsoft Exchange Server 2013 ", "");
									$ExchangeSPLevel = $ExchangeSPLevel.Replace("Microsoft Exchange Server 2016 ", "");
									$ExchangeSPLevel = $ExchangeSPLevel.Replace("Service Pack ", "SP");
									$ExchangeSPLevel = $ExchangeSPLevel.Replace("Cumulative Update ", "CU");
								}
								else
								{
									$ExchangeSPLevel = 0;
								}
							}
							else
							{
								Write-Warning "Cannot detect CU/SP via Remote Registry for $($ExchangeServer.Name)"
							}
						}
						
					}
					# Exchange 2003
					if ($ExchangeMajorVersion -eq 6.5)
					{
						# Mailbox Count
						$MailboxCount = _GetExSvrMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
						# Get Role via WMI
						$tWMI = Get-WMIObject Exchange_Server -Namespace "root\microsoftexchangev2" -Computername $ExchangeServer.Name -Filter "Name='$($ExchangeServer.Name)'"
						if ($tWMI)
						{
							if ($tWMI.IsFrontEndServer) { $Roles = @("FE") }
							else { $Roles = @("BE") }
						}
						else
						{
							Write-Warning "Cannot detect Front End/Back End Server information via WMI for $($ExchangeServer.Name)"
							$Roles += "Unknown"
						}
						# Get Mailbox Statistics using WMI, return in a consistent format
						$tWMI = Get-WMIObject -class Exchange_Mailbox -Namespace ROOT\MicrosoftExchangev2 -ComputerName $ExchangeServer.Name -Filter ("ServerName='$($ExchangeServer.Name)'")
						if ($tWMI)
						{
							$MailboxStatistics = $tWMI | Select @{ Name = "DisplayName"; Expression = { $_.MailboxDisplayName } }, @{ Name = "TotalItemSizeB"; Expression = { $_.Size } }, @{ Name = "TotalDeletedItemSizeB"; Expression = { $_.DeletedMessageSizeExtended } }, @{ Name = "Database"; Expression = { ((get-mailboxdatabase -Identity "$($_.ServerName)\$($_.StorageGroupName)\$($_.StoreName)").identity) } }
						}
						else
						{
							Write-Warning "Cannot retrieve Mailbox Statistics via WMI for $($ExchangeServer.Name)"
							$MailboxStatistics = $null
						}
					}
					# Exchange 2000
					if ($ExchangeMajorVersion -eq "6.0")
					{
						# Mailbox Count
						$MailboxCount = _GetExSvrMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
						# Get Role via ADSI
						$tADSI = [ADSI]"LDAP://$($ExchangeServer.OriginatingServer)/$($ExchangeServer.DistinguishedName)"
						if ($tADSI)
						{
							if ($tADSI.ServerRole -eq 1) { $Roles = @("FE") }
							else { $Roles = @("BE") }
						}
						else
						{
							Write-Warning "Cannot detect Front End/Back End Server information via ADSI for $($ExchangeServer.Name)"
							$Roles += "Unknown"
						}
						$MailboxStatistics = $null
					}
					
					# Return Hashtable
					@{
						Name = $ExchangeServer.Name.ToUpper()
						RealName = $RealName
						ExchangeMajorVersion = $ExchangeMajorVersion
						ExchangeSPLevel = $ExchangeSPLevel
						Edition = $ExchangeServer.Edition
						Mailboxes = $MailboxCount
						OSVersion = $OSVersion;
						OSServicePack = $OSServicePack
						Roles = $Roles
						RollupLevel = $RollupLevel
						RollupVersion = $RollupVersion
						Site = $ExchangeServer.Site.Name
						MailboxStatistics = $MailboxStatistics
						Disks = $Disks
						IntNames = $IntNames
						ExtNames = $ExtNames
						CASArrayName = $CASArrayName
					}
				}
				
				# Sub Function to Get Totals by Version
				function _TotalsByVersion
				{
					param ($ExchangeEnvironment)
					$TotalMailboxesByVersion = @{ }
					if ($ExchangeEnvironment.Sites)
					{
						foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator())
						{
							foreach ($Server in $Site.Value)
							{
								if (!$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"])
								{
									$TotalMailboxesByVersion.Add("$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)", @{ ServerCount = 1; MailboxCount = $Server.Mailboxes })
								}
								else
								{
									$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].ServerCount++
									$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].MailboxCount += $Server.Mailboxes
								}
							}
						}
					}
					if ($ExchangeEnvironment.Pre2007)
					{
						foreach ($FakeSite in $ExchangeEnvironment.Pre2007.GetEnumerator())
						{
							foreach ($Server in $FakeSite.Value)
							{
								if (!$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"])
								{
									$TotalMailboxesByVersion.Add("$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)", @{ ServerCount = 1; MailboxCount = $Server.Mailboxes })
								}
								else
								{
									$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].ServerCount++
									$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].MailboxCount += $Server.Mailboxes
								}
							}
						}
					}
					$TotalMailboxesByVersion
				}
				
				# Sub Function to Get Totals by Role
				function _TotalsByRole
				{
					param ($ExchangeEnvironment)
					# Add Roles We Always Show
					$TotalServersByRole = @{
						"ClientAccess"	   = 0
						"HubTransport"	   = 0
						"UnifiedMessaging" = 0
						"Mailbox"		   = 0
						"Edge"			   = 0
					}
					if ($ExchangeEnvironment.Sites)
					{
						foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator())
						{
							foreach ($Server in $Site.Value)
							{
								foreach ($Role in $Server.Roles)
								{
									if ($TotalServersByRole[$Role] -eq $null)
									{
										$TotalServersByRole.Add($Role, 1)
									}
									else
									{
										$TotalServersByRole[$Role]++
									}
								}
							}
						}
					}
					if ($ExchangeEnvironment.Pre2007["Pre 2007 Servers"])
					{
						
						foreach ($Server in $ExchangeEnvironment.Pre2007["Pre 2007 Servers"])
						{
							
							foreach ($Role in $Server.Roles)
							{
								if ($TotalServersByRole[$Role] -eq $null)
								{
									$TotalServersByRole.Add($Role, 1)
								}
								else
								{
									$TotalServersByRole[$Role]++
								}
							}
						}
					}
					$TotalServersByRole
				}
				
				# Sub Function to return HTML Table for Sites/Pre 2007
				function _GetOverview
				{
					param ($Servers,
						$ExchangeEnvironment,
						$ExRoleStrings,
						$Pre2007 = $False)
					if ($Pre2007)
					{
						$BGColHeader = "#880099"
						$BGColSubHeader = "#8800CC"
						$Prefix = ""
						$IntNamesText = ""
						$ExtNamesText = ""
						$CASArrayText = ""
					}
					else
					{
						$BGColHeader = "#000099"
						$BGColSubHeader = "#0000FF"
						$Prefix = "Site:"
						$IntNamesText = ""
						$ExtNamesText = ""
						$CASArrayText = ""
						$IntNames = @()
						$ExtNames = @()
						$CASArrayName = ""
						foreach ($Server in $Servers.Value)
						{
							$IntNames += $Server.IntNames
							$ExtNames += $Server.ExtNames
							$CASArrayName = $Server.CASArrayName
							
						}
						$IntNames = $IntNames | Sort -Unique
						$ExtNames = $ExtNames | Sort -Unique
						$IntNames = [system.String]::Join(",", $IntNames)
						$ExtNames = [system.String]::Join(",", $ExtNames)
						if ($IntNames)
						{
							$IntNamesText = "Internal Names: $($IntNames)"
							$ExtNamesText = "External Names: $($ExtNames)<br >"
						}
						if ($CASArrayName)
						{
							$CASArrayText = "CAS Array: $($CASArrayName)"
						}
					}
					$Output = "<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Segoe UI,Arial,sans-serif"">
       <col width=""20%""><col width=""20%"">
       <colgroup width=""25%"">";
					
					$ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort Name | %{ $Output += "<col width=""3%"">" }
					$Output += "</colgroup><col width=""20%""><col  width=""20%"">
       <tr bgcolor=""$($BGColHeader)""><th><font color=""#ffffff"">$($Prefix) $($Servers.Key)</font></th>
       <th colspan=""$(($ExchangeEnvironment.TotalServersByRole.Count) + 2)"" align=""left""><font color=""#ffffff"">$($ExtNamesText)$($IntNamesText)</font></th>
       <th align=""center""><font color=""#ffffff"">$($CASArrayText)</font></th></tr>"
					$TotalMailboxes = 0
					$Servers.Value | %{ $TotalMailboxes += $_.Mailboxes }
					$Output += "<tr bgcolor=""$($BGColSubHeader)""><th><font color=""#ffffff"">Mailboxes: $($TotalMailboxes)</font></th><th>"
					$Output += "<font color=""#ffffff"">Exchange Version</font></th>"
					$ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort Name | %{ $Output += "<th><font color=""#ffffff"">$($ExRoleStrings[$_.Key].Short)</font></th>" }
					$Output += "<th><font color=""#ffffff"">OS Version</font></th><th><font color=""#ffffff"">OS Service Pack</font></th></tr>"
					$AlternateRow = 0
					
					foreach ($Server in $Servers.Value)
					{
						$Output += "<tr "
						if ($AlternateRow)
						{
							$Output += " style=""background-color:#dddddd"""
							$AlternateRow = 0
						}
						else
						{
							$AlternateRow = 1
						}
						$Output += "><td>$($Server.Name)"
						if ($Server.RealName -ne $Server.Name)
						{
							$Output += " ($($Server.RealName))"
						}
						$Output += "</td><td>$($ExVersionStrings["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].Long)"
						if ($Server.RollupLevel -gt 0)
						{
							$Output += " UR$($Server.RollupLevel)"
							if ($Server.RollupVersion)
							{
								$Output += " $($Server.RollupVersion)"
							}
						}
						$Output += "</td>"
						$ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort Name | %{
							$Output += "<td"
							if ($Server.Roles -contains $_.Key)
							{
								$Output += " align=""center"" style=""background-color:#00FF00"""
							}
							$Output += ">"
							if (($_.Key -eq "ClusteredMailbox" -or $_.Key -eq "Mailbox" -or $_.Key -eq "BE") -and $Server.Roles -contains $_.Key)
							{
								$Output += $Server.Mailboxes
							}
						}
						
						$Output += "<td>$($Server.OSVersion)</td><td>$($Server.OSServicePack)</td></tr>";
					}
					$Output += "<tr></tr>
       </table><br />"
					$Output
				}
				
				# Sub Function to return HTML Table for Databases
				function _GetDBTable
				{
					param ($Databases)
					# Only Show Archive Mailbox Columns, Backup Columns and Circ Logging if at least one DB has an Archive mailbox, backed up or Cir Log enabled.
					$ShowArchiveDBs = $False
					$ShowLastFullBackup = $False
					$ShowCircularLogging = $False
					$ShowStorageGroups = $False
					$ShowCopies = $False
					$ShowFreeDatabaseSpace = $False
					$ShowFreeLogDiskSpace = $False
					foreach ($Database in $Databases)
					{
						if ($Database.ArchiveMailboxCount -gt 0)
						{
							$ShowArchiveDBs = $True
						}
						if ($Database.LastFullBackup -ne "Not Available")
						{
							$ShowLastFullBackup = $True
						}
						if ($Database.CircularLoggingEnabled -eq "Yes")
						{
							$ShowCircularLogging = $True
						}
						if ($Database.StorageGroup)
						{
							$ShowStorageGroups = $True
						}
						if ($Database.CopyCount -gt 0)
						{
							$ShowCopies = $True
						}
						if ($Database.FreeDatabaseDiskSpace -ne $null)
						{
							$ShowFreeDatabaseSpace = $true
						}
						if ($Database.FreeLogDiskSpace -ne $null)
						{
							$ShowFreeLogDiskSpace = $true
						}
					}
					
					
					$Output = "<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Segoe UI,Arial,sans-serif"">
       
       <tr align=""center"" bgcolor=""#FFD700"">
       <th>Server</th>"
					if ($ShowStorageGroups)
					{
						$Output += "<th>Storage Group</th>"
					}
					$Output += "<th>Database Name</th>
       <th>Mailboxes</th>
       <th>Av. Mailbox Size</th>"
					if ($ShowArchiveDBs)
					{
						$Output += "<th>Archive MBs</th><th>Av. Archive Size</th>"
					}
					$Output += "<th>DB Size</th><th>DB Whitespace</th>"
					if ($ShowFreeDatabaseSpace)
					{
						$Output += "<th>Database Disk Free</th>"
					}
					if ($ShowFreeLogDiskSpace)
					{
						$Output += "<th>Log Disk Free</th>"
					}
					if ($ShowLastFullBackup)
					{
						$Output += "<th>Last Full Backup</th>"
					}
					if ($ShowCircularLogging)
					{
						$Output += "<th>Circular Logging</th>"
					}
					if ($ShowCopies)
					{
						$Output += "<th>Copies (n)</th>"
					}
					
					$Output += "</tr>"
					$AlternateRow = 0;
					foreach ($Database in $Databases)
					{
						$Output += "<tr"
						if ($AlternateRow)
						{
							$Output += " style=""background-color:#dddddd"""
							$AlternateRow = 0
						}
						else
						{
							$AlternateRow = 1
						}
						
						$Output += "><td>$($Database.ActiveOwner)</td>"
						if ($ShowStorageGroups)
						{
							$Output += "<td>$($Database.StorageGroup)</td>"
						}
						$Output += "<td>$($Database.Name)</td>
             <td align=""center"">$($Database.MailboxCount)</td>
             <td align=""center"">$("{0:N2}" -f ($Database.MailboxAverageSize/1MB)) MB</td>"
						if ($ShowArchiveDBs)
						{
							$Output += "<td align=""center"">$($Database.ArchiveMailboxCount)</td> 
                    <td align=""center"">$("{0:N2}" -f ($Database.ArchiveAverageSize/1MB)) MB</td>";
						}
						$Output += "<td align=""center"">$("{0:N2}" -f ($Database.Size/1GB)) GB </td>
             <td align=""center"">$("{0:N2}" -f ($Database.Whitespace/1GB)) GB</td>";
						if ($ShowFreeDatabaseSpace)
						{
							$Output += "<td align=""center"">$("{0:N1}" -f $Database.FreeDatabaseDiskSpace)%</td>"
						}
						if ($ShowFreeLogDiskSpace)
						{
							$Output += "<td align=""center"">$("{0:N1}" -f $Database.FreeLogDiskSpace)%</td>"
						}
						if ($ShowLastFullBackup)
						{
							$Output += "<td align=""center"">$($Database.LastFullBackup)</td>";
						}
						if ($ShowCircularLogging)
						{
							$Output += "<td align=""center"">$($Database.CircularLoggingEnabled)</td>";
						}
						if ($ShowCopies)
						{
							$Output += "<td>$($Database.Copies | %{ $_ }) ($($Database.CopyCount))</td>"
						}
						$Output += "</tr>";
					}
					$Output += "</table><br />"
					
					$Output
				}
				
				
				# Sub Function to neatly update progress
				function _UpProg1
				{
					param ($PercentComplete,
						$Status,
						$Stage)
					$TotalStages = 5
					Write-Progress -id 1 -activity "Get-ExchangeEnvironmentReport" -status $Status -percentComplete (($PercentComplete/$TotalStages) + (1/$TotalStages * $Stage * 100))
				}
				
				# 1. Initial Startup
				
				# 1.0 Check Powershell Version
				if ((Get-Host).Version.Major -eq 1)
				{
					throw "Powershell Version 1 not supported";
				}
				
				# 1.1 Check Exchange Management Shell, attempt to load
				if (!(Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue))
				{
					if (Test-Path "C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1")
					{
						. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'
						Connect-ExchangeServer -auto
					}
					elseif (Test-Path "C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1")
					{
						Add-PSSnapIn Microsoft.Exchange.Management.PowerShell.Admin
						.'C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1'
					}
					else
					{
						throw "Exchange Management Shell cannot be loaded"
					}
				}
				
				# 1.2 Check if -SendMail parameter set and if so check -MailFrom, -MailTo and -MailServer are set
				if ($SendMail)
				{
					if (!$MailFrom -or !$MailTo -or !$MailServer)
					{
						throw "If -SendMail specified, you must also specify -MailFrom, -MailTo and -MailServer"
					}
				}
				
				# 1.3 Check Exchange Management Shell Version
				if ((Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue))
				{
					$E2010 = $false;
					if (Get-ExchangeServer | Where { $_.AdminDisplayVersion.Major -gt 14 })
					{
						Write-Warning "Exchange 2010 or higher detected. You'll get better results if you run this script from the latest management shell"
					}
				}
				else
				{
					
					$E2010 = $true
					$localserver = get-exchangeserver $Env:computername
					$localversion = $localserver.admindisplayversion.major
					if ($localversion -eq 15) { $E2013 = $true }
					
				}
				
				# 1.4 Check view entire forest if set (by default, true)
				if ($E2010)
				{
					Set-ADServerSettings -ViewEntireForest:$ViewEntireForest
				}
				else
				{
					$global:AdminSessionADSettings.ViewEntireForest = $ViewEntireForest
				}
				
				# 1.5 Initial Variables
				
				# 1.5.1 Hashtable to update with environment data
				$ExchangeEnvironment = @{
					Sites = @{ }
					Pre2007 = @{ }
					Servers = @{ }
					DAGs  = @()
					NonDAGDatabases = @()
				}
				# 1.5.7 Exchange Major Version String Mapping
				$ExMajorVersionStrings = @{
					"6.0" = @{ Long = "Exchange 2000"; Short = "E2000" }
					"6.5" = @{ Long = "Exchange 2003"; Short = "E2003" }
					"8"   = @{ Long = "Exchange 2007"; Short = "E2007" }
					"14"  = @{ Long = "Exchange 2010"; Short = "E2010" }
					"15"  = @{ Long = "Exchange 2013"; Short = "E2013" }
					"15.1" = @{ Long = "Exchange 2016"; Short = "E2016" }
				}
				# 1.5.8 Exchange Service Pack String Mapping
				$ExSPLevelStrings = @{
					"0"   = "RTM"
					"1"   = "SP1"
					"2"   = "SP2"
					"3"   = "SP3"
					"4"   = "SP4"
					"SP1" = "SP1"
					"SP2" = "SP2"
				}
				# Add many CUs               
				for ($i = 1; $i -le 20; $i++)
				{
					$ExSPLevelStrings.Add("CU$($i)", "CU$($i)");
				}
				# 1.5.9 Populate Full Mapping using above info
				$ExVersionStrings = @{ }
				foreach ($Major in $ExMajorVersionStrings.GetEnumerator())
				{
					foreach ($Minor in $ExSPLevelStrings.GetEnumerator())
					{
						$ExVersionStrings.Add("$($Major.Key).$($Minor.Key)", @{ Long = "$($Major.Value.Long) $($Minor.Value)"; Short = "$($Major.Value.Short)$($Minor.Value)" })
					}
				}
				# 1.5.10 Exchange Role String Mapping
				$ExRoleStrings = @{
					"ClusteredMailbox" = @{ Short = "ClusMBX"; Long = "CCR/SCC Clustered Mailbox" }
					"Mailbox"		   = @{ Short = "MBX"; Long = "Mailbox" }
					"ClientAccess"	   = @{ Short = "CAS"; Long = "Client Access" }
					"HubTransport"	   = @{ Short = "HUB"; Long = "Hub Transport" }
					"UnifiedMessaging" = @{ Short = "UM"; Long = "Unified Messaging" }
					"Edge"			   = @{ Short = "EDGE"; Long = "Edge Transport" }
					"FE"			   = @{ Short = "FE"; Long = "Front End" }
					"BE"			   = @{ Short = "BE"; Long = "Back End" }
					"Hybrid"		   = @{ Short = "HYB"; Long = "Hybrid" }
					"Unknown"		   = @{ Short = "Unknown"; Long = "Unknown" }
				}
				
				# 2 Get Relevant Exchange Information Up-Front
				
				# 2.1 Get Server, Exchange and Mailbox Information
				_UpProg1 1 "Getting Exchange Server List" 1
				$ExchangeServers = [array](Get-ExchangeServer $ServerFilter)
				if (!$ExchangeServers)
				{
					throw "No Exchange Servers matched by -ServerFilter ""$($ServerFilter)"""
				}
				$HybridServers = @()
				if (Get-Command Get-HybridConfiguration -ErrorAction SilentlyContinue)
				{
					$HybridConfig = Get-HybridConfiguration
					$HybridConfig.ReceivingTransportServers | %{ $HybridServers += $_.Name }
					$HybridConfig.SendingTransportServers | %{ $HybridServers += $_.Name }
					$HybridServers = $HybridServers | Sort-Object -Unique
				}
				
				_UpProg1 10 "Getting Mailboxes" 1
				$Mailboxes = [array](Get-Mailbox -ResultSize Unlimited) | Where { $_.Server -like $ServerFilter }
				if ($E2010)
				{
					_UpProg1 60 "Getting Archive Mailboxes" 1
					$ArchiveMailboxes = [array](Get-Mailbox -Archive -ResultSize Unlimited) | Where { $_.Server -like $ServerFilter }
					_UpProg1 70 "Getting Remote Mailboxes" 1
					$RemoteMailboxes = [array](Get-RemoteMailbox -ResultSize Unlimited)
					$ExchangeEnvironment.Add("RemoteMailboxes", $RemoteMailboxes.Count)
					_UpProg1 90 "Getting Databases" 1
					if ($E2013)
					{
						$Databases = [array](Get-MailboxDatabase -IncludePreExchange2013 -Status) | Where { $_.Server -like $ServerFilter }
					}
					elseif ($E2010)
					{
						$Databases = [array](Get-MailboxDatabase -IncludePreExchange2010 -Status) | Where { $_.Server -like $ServerFilter }
					}
					$DAGs = [array](Get-DatabaseAvailabilityGroup) | Where { $_.Servers -like $ServerFilter }
				}
				else
				{
					$ArchiveMailboxes = $null
					$ArchiveMailboxStats = $null
					$DAGs = $null
					_UpProg1 90 "Getting Databases" 1
					$Databases = [array](Get-MailboxDatabase -IncludePreExchange2007 -Status) | Where { $_.Server -like $ServerFilter }
					$ExchangeEnvironment.Add("RemoteMailboxes", 0)
				}
				
				# 2.3 Populate Information we know
				$ExchangeEnvironment.Add("TotalMailboxes", $Mailboxes.Count + $ExchangeEnvironment.RemoteMailboxes);
				
				# 3 Process High-Level Exchange Information
				
				# 3.1 Collect Exchange Server Information
				for ($i = 0; $i -lt $ExchangeServers.Count; $i++)
				{
					_UpProg1 ($i/$ExchangeServers.Count * 100) "Getting Exchange Server Information" 2
					# Get Exchange Info
					$ExSvr = _GetExSvr -E2010 $E2010 -ExchangeServer $ExchangeServers[$i] -Mailboxes $Mailboxes -Databases $Databases -Hybrids $HybridServers
					# Add to site or pre-Exchange 2007 list
					if ($ExSvr.Site)
					{
						# Exchange 2007 or higher
						if (!$ExchangeEnvironment.Sites[$ExSvr.Site])
						{
							$ExchangeEnvironment.Sites.Add($ExSvr.Site, @($ExSvr))
						}
						else
						{
							$ExchangeEnvironment.Sites[$ExSvr.Site] += $ExSvr
						}
					}
					else
					{
						# Exchange 2003 or lower
						if (!$ExchangeEnvironment.Pre2007["Pre 2007 Servers"])
						{
							$ExchangeEnvironment.Pre2007.Add("Pre 2007 Servers", @($ExSvr))
						}
						else
						{
							$ExchangeEnvironment.Pre2007["Pre 2007 Servers"] += $ExSvr
						}
					}
					# Add to Servers List
					$ExchangeEnvironment.Servers.Add($ExSvr.Name, $ExSvr)
				}
				
				# 3.2 Calculate Environment Totals for Version/Role using collected data
				_UpProg1 1 "Getting Totals" 3
				$ExchangeEnvironment.Add("TotalMailboxesByVersion", (_TotalsByVersion -ExchangeEnvironment $ExchangeEnvironment))
				$ExchangeEnvironment.Add("TotalServersByRole", (_TotalsByRole -ExchangeEnvironment $ExchangeEnvironment))
				
				# 3.4 Populate Environment DAGs
				_UpProg1 5 "Getting DAG Info" 3
				if ($DAGs)
				{
					foreach ($DAG in $DAGs)
					{
						$ExchangeEnvironment.DAGs += (_GetDAG -DAG $DAG)
					}
				}
				
				# 3.5 Get Database information
				_UpProg1 60 "Getting Database Info" 3
				for ($i = 0; $i -lt $Databases.Count; $i++)
				{
					$Database = _GetDB -Database $Databases[$i] -ExchangeEnvironment $ExchangeEnvironment -Mailboxes $Mailboxes -ArchiveMailboxes $ArchiveMailboxes -E2010 $E2010
					$DAGDB = $false
					for ($j = 0; $j -lt $ExchangeEnvironment.DAGs.Count; $j++)
					{
						if ($ExchangeEnvironment.DAGs[$j].Members -contains $Database.ActiveOwner)
						{
							$DAGDB = $true
							$ExchangeEnvironment.DAGs[$j].Databases += $Database
						}
					}
					if (!$DAGDB)
					{
						$ExchangeEnvironment.NonDAGDatabases += $Database
					}
					
					
				}
				
				# 4 Write Information
				_UpProg1 5 "Writing HTML Report Header" 4
				# Header
				$Output = "<html>
<body>
<font size=""1"" face=""Segoe UI,Arial,sans-serif"">
<h2 align=""center"">Exchange Environment Report</h3>
<h4 align=""center"">Generated $((Get-Date).ToString())</h5>
</font>
<table border=""0"" cellpadding=""3"" style=""font-size:8pt;font-family:Segoe UI,Arial,sans-serif"">
<tr bgcolor=""#009900"">
<th colspan=""$($ExchangeEnvironment.TotalMailboxesByVersion.Count)""><font color=""#ffffff"">Total Servers:</font></th>"
				if ($ExchangeEnvironment.RemoteMailboxes)
				{
					$Output += "<th colspan=""$($ExchangeEnvironment.TotalMailboxesByVersion.Count + 2)""><font color=""#ffffff"">Total Mailboxes:</font></th>"
				}
				else
				{
					$Output += "<th colspan=""$($ExchangeEnvironment.TotalMailboxesByVersion.Count + 1)""><font color=""#ffffff"">Total Mailboxes:</font></th>"
				}
				$Output += "<th colspan=""$($ExchangeEnvironment.TotalServersByRole.Count)""><font color=""#ffffff"">Total Roles:</font></th></tr>
<tr bgcolor=""#00CC00"">"
				# Show Column Headings based on the Exchange versions we have
				$ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort Name | %{ $Output += "<th>$($ExVersionStrings[$_.Key].Short)</th>" }
				$ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort Name | %{ $Output += "<th>$($ExVersionStrings[$_.Key].Short)</th>" }
				if ($ExchangeEnvironment.RemoteMailboxes)
				{
					$Output += "<th>Office 365</th>"
				}
				$Output += "<th>Org</th>"
				$ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort Name | %{ $Output += "<th>$($ExRoleStrings[$_.Key].Short)</th>" }
				$Output += "<tr>"
				$Output += "<tr align=""center"" bgcolor=""#dddddd"">"
				$ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort Name | %{ $Output += "<td>$($_.Value.ServerCount)</td>" }
				$ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort Name | %{ $Output += "<td>$($_.Value.MailboxCount)</td>" }
				if ($RemoteMailboxes)
				{
					$Output += "<th>$($ExchangeEnvironment.RemoteMailboxes)</th>"
				}
				$Output += "<td>$($ExchangeEnvironment.TotalMailboxes)</td>"
				$ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort Name | %{ $Output += "<td>$($_.Value)</td>" }
				$Output += "</tr><tr><tr></table><br>"
				
				# Sites and Servers
				_UpProg1 20 "Writing HTML Site Information" 4
				foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator())
				{
					$Output += _GetOverview -Servers $Site -ExchangeEnvironment $ExchangeEnvironment -ExRoleStrings $ExRoleStrings
				}
				_UpProg1 40 "Writing HTML Pre-2007 Information" 4
				foreach ($FakeSite in $ExchangeEnvironment.Pre2007.GetEnumerator())
				{
					$Output += _GetOverview -Servers $FakeSite -ExchangeEnvironment $ExchangeEnvironment -ExRoleStrings $ExRoleStrings -Pre2007:$true
				}
				
				_UpProg1 60 "Writing HTML DAG Information" 4
				foreach ($DAG in $ExchangeEnvironment.DAGs)
				{
					if ($DAG.MemberCount -gt 0)
					{
						# Database Availability Group Header
						$Output += "<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Segoe UI,Arial,sans-serif"">
             <col width=""20%""><col width=""10%""><col width=""70%"">
             <tr align=""center"" bgcolor=""#FF8000 ""><th>Database Availability Group Name</th><th>Member Count</th>
             <th>Database Availability Group Members</th></tr>
             <tr><td>$($DAG.Name)</td><td align=""center"">
             $($DAG.MemberCount)</td><td>"
						$DAG.Members | % { $Output += "$($_) " }
						$Output += "</td></tr></table>"
						
						# Get Table HTML
						$Output += _GetDBTable -Databases $DAG.Databases
					}
					
				}
				
				if ($ExchangeEnvironment.NonDAGDatabases.Count)
				{
					_UpProg1 80 "Writing HTML Non-DAG Database Information" 4
					$Output += "<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Segoe UI,Arial,sans-serif"">
         <tr bgcolor=""#FF8000""><th>Mailbox Databases (Non-DAG)</th></table>"
					$Output += _GetDBTable -Databases $ExchangeEnvironment.NonDAGDatabases
				}
				
				
				# End
				_UpProg1 90 "Finishing off.." 4
				$Output += "</body></html>";
				$Output | Out-File $HTMLReport
				
				
				if ($SendMail)
				{
					_UpProg1 95 "Sending mail message.." 4
					Send-MailMessage -Attachments $HTMLReport -To $MailTo -From $MailFrom -Subject "Exchange Environment Report" -BodyAsHtml $Output -SmtpServer $MailServer
				}
				
				$Reboot = $true
			}
			#endregion   
			#region Option 62) Generate Mailbox Size and Information Reports
			62 {
				"This function is not yet implemented"
				#      Generate Mailbox Size and Information Reports
				<#empty
				$Reboot = $false#>
			}
			#endregion   
			#region Option 63) Generate Reports for Exchange ActiveSync Device Statistics
			63 {
				generateEASDeviceStats
			}
			#endregion   
			#region Option 64) Exchange Analyzer
			64 {
				"This function is not yet implemented"
				#      Exchange Analyzer
				<#empty
				$Reboot = $false#>
			}
			#endregion   
			#region Option 65) Generate Report Total Emails Sent and Received Per Day and Size
			65 {
				#      Generate Report Total Emails Sent and Received Per Day and Size
				# Script:    TotalEmailsSentReceivedPerDay.ps1
				# Purpose:   Get the number of e-mails sent and received per day
				# Author:    Nuno Mota
				# Date:             October 2010
				#region user input
				"Get the number of e-mails sent and received per day"
				"Enter start date"
				[INT]$MM = Read-Host "Month"
				[INT]$DD = Read-Host "Day"
				[INT]$YY = Read-Host "Year"
				
				[INT]$noOfdays = Read-Host "Enter number of days. (Start date included)"
				#endregion
				
				[Int64]$intSent = $intRec = 0
				[Int64]$intSentSize = $intRecSize = 0
				[String]$strEmails = $null
				
				Write-Host "DayOfWeek,Date,Sent,Sent Size (MB),Received,Received Size (MB)" -ForegroundColor Yellow
				
				Do
				{
					# Start building the variable that will hold the information for the day 
					$strEmails = "$($From.DayOfWeek),$($From.ToShortDateString()),"
					
					$intSent = $intRec = 0
					(Get-TransportService) | Get-MessageTrackingLog -ResultSize Unlimited -Start $From -End $To | ForEach {
						# Sent E-mails 
						If ($_.EventId -eq "RECEIVE" -and $_.Source -eq "STOREDRIVER")
						{
							$intSent++
							$intSentSize += $_.TotalBytes
						}
						
						# Received E-mails 
						If ($_.EventId -eq "DELIVER")
						{
							$intRec += $_.RecipientCount
							$intRecSize += $_.TotalBytes
						}
					}
					
					$intSentSize = [Math]::Round($intSentSize/1MB, 0)
					$intRecSize = [Math]::Round($intRecSize/1MB, 0)
					
					# Add the numbers to the $strEmails variable and print the result for the day 
					$strEmails += "$intSent,$intSentSize,$intRec,$intRecSize"
					$strEmails
					
					# Increment the From and To by one day 
					$From = $From.AddDays(1)
					$To = $From.AddDays(1)
				}
				While ($To -lt (Get-Date))
				#While ($To -lt (Get-Date "01/12/2011"))
				
				$Reboot = $false
				
			}
			#endregion   
			#region Option 66) Generate HTML Report for Mailbox Permissions
			66 {
				#      Generate HTML Report for Mailbox Permissions
                           <#
    .SYNOPSIS
    Creates a HTML Report showing Sendas, Full Access and Send on Behalf Permission Information for Each Mailbox for your Exchange Organization, selected database or for a single user.
   
       Serkan Varoglu
       
       http:\\Mshowto.org
       http:\\Get-Mailbox.org
       
       THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
       RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
       
       Version 1.1, 5 March 2012
       
    .DESCRIPTION
       
    Creates a HTML Report showing Sendas, Full Access and Send on Behalf Permission Information for Each Mailbox for your Exchange Organization, selected database or for a single user.
       By Default Inherited Send As permission and NT Authority\Self account will not be shown in the report unless you run the script with the parameters listed below.
       Also by default all mailboxes will be reported if you want to report a single database, you can use -database parameter to specify your database name or you can get the report for a single user.
       
       .PARAMETER HTMLReport
    Filename to write HTML Report to
       
       .PARAMETER Database
    By default this script will report all mailboxes. If you want to report mailboxes in a single database, you can use this parameter to input your database name.
       
       .PARAMETER Mailbox
    By default this script will report all mailboxes. If you want to report a single mailbox, you can use this parameter to input the mailbox you want to report.
       
       .SWITCH ShowInherited
       If ShowInherited is added as switch the report will show Inherited Sendas permissions for mailboxes as well.
       
       .SWITCH ShowSelf
       If ShowSelf is added as switch the report will show "NT Authority\Self" sendas permission for mailboxes as well.
       
       .EXAMPLE
    Generate the HTML report 
    .\Report-Permissions.ps1 -HTMLReport "C:\Users\SVaroglu\Desktop\MailboxPermissionReport.HTML"
       
#>
				
				#region user input
				$HTMLReport = Read-Host "Specifiy alternate path and name for report file. Default is ""mailboxpermissionsreport.html"""
				if ($HTMLReport = "")
				{
					$ReportFile = "mailboxpermissionsreport.html"
				}
				$ShowInheritedYN = Read-Host "List inherited SendAs and Full Access permissions?[Y/N] Default is [Y]"
				switch ($ShowInheritedYN)
				{
					Y{ $ShowInherited = $true }
					N{ $ShowInherited = $false }
					default { $ShowInherited = $true }
				}
				$ShowSelfYN = Read-Host "List NT Authority\Self Permission ?[Y/N] Default is [Y]"
				switch ($ShowInheritedYN)
				{
					Y{ $ShowSelf = $true }
					N{ $ShowSelf = $false }
					default { $ShowSelf = $true }
				}
				$MailboxYN = Read-Host "Specify a mailbox to report?[Y/N] Default is [N]"
				switch ($MailboxYN)
				{
					Y{ $Mailbox = Read-Host "Enter mailbox name" }
					N{ $Mailbox = $null }
					default { $Mailbox = $null }
				}
				#endregion
				$Watch = [System.Diagnostics.Stopwatch]::StartNew()
				$WarningPreference = "SilentlyContinue"
				$ErrorActionPreference = "SilentlyContinue"
				$ShowInherited = $ShowInherited.IsPresent
				$ShowSelf = $ShowSelf.IsPresent
				$u = 1
				$s = 0
				$f = 0
				$b = 0
				$n = 0
				$nj = -1
				$gj = -1
				if (!$database) { $dbnull = 0 }
				if (!$mailbox) { $mbnull = 0 }
				if (($ShowSelf -like "false") -and ($ShowInherited -like "false"))
				{ $gentitle = "Mailboxes With Custom Permissions" }
				else
				{ $gentitle = "Mailboxes" }
				$gen = "<table border=""1"" bordercolor=""#4384D3"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#4384D3"" align=""center""><th colspan=""5""><font color=""#FFFFFF"">$($gentitle)</font></th></tr><tr>"
				$inh = "<table border=""1"" bordercolor=""#4384D3"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#4384D3"" align=""center""><th colspan=""5""><font color=""#FFFFFF"">Mailboxes With Only Inherited Permissions</font></th></tr><tr>"
				function _Progress
				{
					param ($PercentComplete,
						$Status)
					Write-Progress -id 1 -activity "Report for Mailboxes" -status $Status -percentComplete ($PercentComplete)
				}
				_Progress (($u * 100)/100) "Collecting Mailbox Information"
				if (!$database -and !$mailbox)
				{
					$mailboxes = get-mailbox -resultsize unlimited | Sort-Object name
				}
				elseif ($database -and !$mailbox)
				{
					$mailboxes = get-mailbox -database $database -resultsize unlimited | Sort-Object name
				}
				elseif (!$database -and $mailbox)
				{
					$mailboxes = get-mailbox $mailbox
				}
				else
				{
					Write-Host -ForegroundColor Cyan "Please choose database or single mailbox. Both Parameters can not be used at the same time. Ended without compiling a report."
					exit
				}
				$mcount = ($mailboxes | measure-object).count
				if ($mcount -eq 0)
				{
					Write-Host -ForegroundColor Cyan "No Mailbox Found. Ended without compiling a report. Please Check Your Input."
					exit
				}
				foreach ($mailbox in $mailboxes)
				{
					_Progress (($u * 95)/$mcount) "Processing $mailbox, $($u) of $($mcount) Mailboxes."
					$SenderBody = ""
					$FullBody = ""
					$BehalfBody = ""
					$sendbehalfs = Get-Mailbox $mailbox | select-object -expand grantsendonbehalfto | select-object -expand rdn | Sort-Object Unescapedname
					if (($ShowSelf -like "true") -and ($ShowInherited -like "true"))
					{
						$senders = Get-ADPermission $mailbox.identity | ?{ ($_.extendedrights -like "*send-as*") } | Sort-Object name
						$fullsenders = Get-Mailbox $mailbox | Get-MailboxPermission | ?{ ($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") }
					}
					elseif (($ShowSelf -like "false") -and ($ShowInherited -like "true"))
					{
						$senders = Get-ADPermission $mailbox.identity | ?{ ($_.extendedrights -like "*send-as*") -and ($_.User -notlike "NT Authority\self") } | Sort-Object name
						$fullsenders = Get-Mailbox $mailbox | Get-MailboxPermission | ?{ ($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\self*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") }
					}
					elseif (($ShowSelf -like "true") -and ($ShowInherited -like "false"))
					{
						$senders = Get-ADPermission $mailbox.identity | ?{ ($_.extendedrights -like "*send-as*") -and ($_.isinherited -like "false") } | Sort-Object name
						$fullsenders = Get-Mailbox $mailbox | Get-MailboxPermission | ?{ ($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") -and ($_.IsInherited -like "false") }
					}
					else
					{
						$senders = Get-ADPermission $mailbox.identity | ?{ ($_.extendedrights -like "*send-as*") -and ($_.isinherited -like "false") -and ($_.User -notlike "NT Authority\self") } | Sort-Object name
						$fullsenders = Get-Mailbox $mailbox | Get-MailboxPermission | ?{ ($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\self*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") -and ($_.IsInherited -like "false") }
					}
					if (!$senders -and !$fullsenders -and !$sendbehalfs)
					{
						$n++
						if ($nj -eq 4)
						{
							$inh += "</tr><tr><td>$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</td>"
							$nj = 0
						}
						else
						{
							$inh += "<td>$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</td>"
							$nj++
						}
					}
					else
					{
						if ($gj -eq 4)
						{
							$gen += "</tr><tr><td><a href=""#$($mailbox.Name)"">$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</a></td>"
							$gj = 0
						}
						else
						{
							$gen += "<td><a href=""#$($mailbox.Name)"">$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</a></td>"
							$gj++
						}
						$MailboxTable = "<table border=""1"" bordercolor=""#1F497B"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#1F497B"" align=""center""><th colspan=""3"" ><font color=""#FFFFFF""><a name=""$($mailbox.Name)"">$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</font></a></th></tr><tr>"
						if (!$senders)
						{
							$SenderBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#696969"" align=""center"" valign=""top""><font color=""#FFFF99"">No User has Send As Permission On This Mailbox</font></td></table></td>"
						}
						else
						{
							$SenderBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><tr><td colspan=""3"" align=""center"" valign=""top"" bgcolor=""#696969""><font color=""#FFFFE0"">Send-As Permissions</font></td></tr><tr bgcolor=""#878787"" align=""center"">
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Send as Permission Owner</font></td>
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Deny</font></td>
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Inherited</font></td>
                                        </tr>"
							foreach ($sender in $senders)
							{
								if (0, 2, 4, 6, 8 -contains "$sj"[-1] - 48)
								{
									$bgcolor = "'#E8E8E8'"
								}
								else
								{
									$bgcolor = "'#C8C8C8'"
								}
								$SenderBody += "<tr align=""center"" bgcolor=$($bgcolor)>"
								$SenderBody += "<td><font color=""#003333"">$($sender.user)</font></td>"
								if ($sender.deny -like "true") { $font = "red" }
								else { $font = "'#000000'" }
								$SenderBody += "<td><font color=$font>$($sender.deny)</font></td>"
								if ($sender.isinherited -like "false") { $font = "red" }
								else { $font = "'#000000'" }
								$SenderBody += "<td><font color=$font>$($sender.isinherited)</font></td>"
								$SenderBody += "</tr>"
								$sj++
							}
							$SenderBody += "</table></td>"
							$s++
						}
						
						if (!$fullsenders)
						{
							$FullBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#696969"" align=""center"" valign=""top""><font color=""#FFFF99"">No User has Full Access On This Mailbox</font></td></table></td>"
						}
						else
						{
							$FullBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><tr><td colspan=""3"" align=""center"" valign=""top"" bgcolor=""#696969""><font color=""#FFFFE0"">Full Access Permissions</font></td></tr><tr bgcolor=""#878787"" align=""center"">
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Full Access Permission Owner</font></td>
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Deny</font></td>
                                        <td nowrap=""nowrap""><font color=""#FFFFFF"">Inherited</font></td>
                                        </tr>"
							foreach ($fullsender in $fullsenders)
							{
								if (0, 2, 4, 6, 8 -contains "$fj"[-1] - 48)
								{
									$bgcolor = "'#E8E8E8'"
								}
								else
								{
									$bgcolor = "'#C8C8C8'"
								}
								$FullBody += "<tr align=""center"" bgcolor=$($bgcolor)>"
								$FullBody += "<td><font color=""#003333"">$($fullsender.user)</font></td>"
								if ($fullsender.deny -like "true") { $font = "red" }
								else { $font = "'#000000'" }
								$FullBody += "<td><font color=$font>$($fullsender.deny)</font></td>"
								if ($fullsender.isinherited -like "false") { $font = "red" }
								else { $font = "'#000000'" }
								$FullBody += "<td><font color=$font>$($fullsender.isinherited)</font></td>"
								$FullBody += "</tr>"
								$fj++
							}
							$FullBody += "</table></td>"
							$f++
						}
						
						if (!$sendbehalfs)
						{
							$BehalfBody += "<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#696969"" align=""center"" valign=""top""><font color=""#FFFF99"">No User has Send on Behalf On This Mailbox</font></td></table></td>"
						}
						else
						{
							$BehalfBody += "<td align=""center"" valign=""top"" width=""33%"">
                                        <table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
                                        <tr><td align=""center valign=""top"" bgcolor=""#696969""><font color=""#FFFFE0"">Send on Behalf</font></td></tr>
                                        <tr><td bgcolor=""#878787"" nowrap=""nowrap""><font color=""#FFFFFF"">Send On Behalf Permission Owner</font></td></tr>"
							foreach ($sendbehalf in $sendbehalfs)
							{
								if (0, 2, 4, 6, 8 -contains "$bj"[-1] - 48)
								{
									$bgcolor = "'#E8E8E8'"
								}
								else
								{
									$bgcolor = "'#C8C8C8'"
								}
								$BehalfBody += "<tr align=""center"" bgcolor=$($bgcolor)>"
								$BehalfBody += "<td><font color=""#003333"">$($sendbehalf.unescapedname)</font></td>"
								$BehalfBody += "</tr>"
								$bj++
							}
							$BehalfBody += "</table></td>"
							$b++
						}
						$Table += $MailboxTable + $SenderBody + $FullBody + $BehalfBody + "</tr></table><br><a href=""#top"">&#9650;</a><hr /><br>"
					}
					$u++
				}
				_Progress (98) "Completing"
				if (($ShowSelf -like "false") -and ($ShowInherited -like "false"))
				{
					if (($dbnull -eq 0) -and ($mbnull -eq 0))
					{
						$Summary = "<table style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#FFE87C"" >In your Exchange Organization there are $($mcount) mailboxes present."
						$Summary += "Send as Permission explicity configured on $($s) of these mailboxes. Full Access Permission explicity configured on $($f) of these mailboxes. Send on Behalf explicity configured on $($b) of these mailboxes and $($n) mailbox has inherited permissions only.<td></table><br>"
					}
					elseif (($dbnull -ne 0) -and ($mbnull -eq 0))
					{
						$Summary = "<table style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#FFE87C"" >In $($database) mailbox database, there are $($mcount) mailboxes present."
						$Summary += "Send as Permission explicity configured on $($s) of these mailboxes. Full Access Permission explicity configured on $($f) of these mailboxes. Send on Behalf explicity configured on $($b) of these mailboxes and $($n) mailbox has inherited permissions only.<td></table><br>"
					}
					$Header = "
       <body>
       <font size=""1"" face=""Arial,sans-serif"">
       <h3 align=""center"">Mailbox Send As, Full Permission and Send on Behalf Report</h3>
       <h4 align=""center"">Generated $((Get-Date).ToString())</h4>"
					$inh += "</tr></table><br>"
					$gen += "</tr></table><br>"
					$Footer = "</table></center><br><br>
       Scripted by <a href=""http://www.get-mailbox.org"">Serkan Varoglu</a>.  
       Elapsed Time To Complete This Report: $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString())</body></html>"
					if (($dbnull -eq 0) -and ($mbnull -eq 0))
					{
						$Output = $Header + $Summary + $gen + $inh + "<br><hr /><br>" + $Table + $Footer
					}
					elseif (($dbnull -ne 0) -and ($mbnull -eq 0))
					{
						$Output = $Header + $Summary + $gen + $inh + "<br><hr /><br>" + $Table + $Footer
					}
					else
					{
						if (($s -eq 0) -and ($f -eq 0) -and ($b -eq 0))
						{
							$Note = "<center></font><b>Mailbox for $($Mailbox.name) ( $($Mailbox.primarysmtpaddress) ), does not have any explicit permissions set for Send As, Full Access or Send on Behalf</b></center>"
						}
						$Output = $Header + $Note + $Table + $Footer
					}
				}
				else
				{
					$Header = "
       <body>
       <font size=""1"" face=""Arial,sans-serif"">
       <h3 align=""center"">Mailbox Send As, Full Permission and Send on Behalf Report</h3>
       <a name=""top""><h4 align=""center"">Generated $((Get-Date).ToString())</h4></a>
       "
					$inh += "</tr></table><br>"
					$gen += "</tr></table><br>"
					$Footer = "</table></center><br><br>
       <font size=""1"" face=""Arial,sans-serif"">Scripted by <a href=""http://www.get-mailbox.org"">Serkan Varoglu</a>.  
       Elapsed Time To Complete This Report: $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString())</font></body></html>"
					$Output = $Header + $gen + $Table + $Footer
					
				}
				$Output | Out-File $HTMLReport
				
				$Reboot = $false
				
			}
			#endregion   
			#region Option 70) Export Office 365 User Last Logon Date to CSV File
			70 {
				if ($O365Creds -eq $null)
				{
					$O365Creds = Get-Credential -Message "Enter your 0365 credentials!"
				}
				O365ExportLastLogonDate
				$Reboot = $false
				
			}
			#endregion   
			#region Option 71) List all Distribution Groups and their Membership in Office 365
			71 {
				if ($O365Creds -eq $null)
				{
					$O365Creds = Get-Credential -Message "Enter your 0365 credentials!"
				}
				O365ListDistGroupsAndMemberships -O365Creds $O365Creds
			}
			#endregion   
			#region Option 72) Office 365 Mail Traffic Statistics by User
			72 {
				if ($O365Creds -eq $null)
				{
					$O365Creds = Get-Credential -Message "Enter your 0365 credentials!"
				}
				O365MailTrafficStatsbyUser -O365Creds $O365Creds
			}
			#endregion   
			#region Option 73) Export a Licence reconciliation report from Office 365
			73 {
				if ($O365Creds -eq $null)
				{
					$O365Creds = Get-Credential -Message "Enter your 0365 credentials!"
				}
				O365ExportLicenseReconcilation ($O365Creds)
			}
			#endregion   
			#region Option 74) Export mailbox permissions from Office 365 to CSV file
			74 {
				#region user input
				$HTMLReport = Read-Host "Specifiy alternate path and name for report file. Default is ""mailboxpermissionsreport.html"""
				if ($HTMLReport = "")
				{
					$HTMLReport = "O365MailboxFolderPermissionsReport.html"
				}
				$ShowInheritedYN = Read-Host "List inherited SendAs and Full Access permissions?[Y/N] Default is [Y]"
				switch ($ShowInheritedYN)
				{
					Y{ $ShowInherited = $true }
					N{ $ShowInherited = $false }
					default { $ShowInherited = $true }
				}
				$ShowSelfYN = Read-Host "List NT Authority\Self Permission ?[Y/N] Default is [Y]"
				switch ($ShowInheritedYN)
				{
					Y{ $ShowSelf = $true }
					N{ $ShowSelf = $false }
					default { $ShowSelf = $true }
				}
				$MailboxYN = Read-Host "Specify a single mailbox only?[Y/N] Default is [N]"
				switch ($MailboxYN)
				{
					Y{ $Mailbox = Read-Host "Enter mailbox name" }
					N{ $Mailbox = $null }
					default { $Mailbox = $null }
				}
				if ($Mailbox -eq $null)
				{
					$dbYN = Read-Host "Specify a single database only?[Y/N] Default is [N]"
					switch ($dbYN)
					{
						Y{ $Database = Read-Host "Enter database name" }
						N{ $Database = $null }
						default { $Database = $null }
					}
				}
				if ($O365Creds -eq $null)
				{
					$O365Creds = Get-Credential -Message "Enter your 0365 credentials!"
				}
				#endregion
				
				ExportMBXFolderPermissions -HTMLReport $HTMLReport -ShowInherited:$ShowInherited -ShowSelf:$ShowSelf -Database $Database -Mailbox $Mailbox -O365 -O365Creds $O365Creds
			}
			#endregion   
						#region Option 75) Microsoft 365 Mailboxes with Synchronized Mobile Devices - by Tony Redmond
			75 {
				# An example script to show how to extract mobile device statistics from devices registred with Exchange Online mailboxes
                # https://github.com/12Knocksinna/Office365itpros/blob/master/Report-MobileDevices.PS1

        $directory23 = "C:\mdm\"

if (-not (Test-Path -Path $directory23 -PathType Container)) {
    New-Item -Path $directory23 -ItemType Directory
}
        
        $HtmlHead ="<html>
	    <style>
	    BODY{font-family: Arial; font-size: 8pt;}
	    H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	    H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	    H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	    TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	    TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	    TD{border: 1px solid #969595; padding: 5px; }
	    td.pass{background: #B7EB83;}
	    td.warn{background: #FFF275;}
	    td.fail{background: #FF2626; color: #ffffff;}
	    td.info{background: #85D4FF;}
	    </style>
	    <body>
           <div align=center>
           <p><h1>Microsoft 365 Mailboxes with Synchronized Mobile Devices</h1></p>
           <p><h3>Generated: " + (Get-Date -format 'dd-MMM-yyyy hh:mm tt') + "</h3></p></div>"

$Version = "1.0"
$HtmlReportFile = "C:\mdm\MobileDevices.html"
$CSVReportFile = "C:\mdm\MobileDevices.csv"

Connect-ExchangeOnline

$Organization = Get-OrganizationConfig | Select-Object -ExpandProperty DisplayName
[array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Sort-Object DisplayName
If (!($Mbx)) { Write-Host "Unable to find any user mailboxes..." ; break }

$Report = [System.Collections.Generic.List[Object]]::new() 

[int]$i = 0
ForEach ($M in $Mbx) {
 $i++
 Write-Host ("Scanning mailbox {0} for registered mobile devices... {1}/{2}" -f $M.DisplayName, $i, $Mbx.count)
 [array]$Devices = Get-MobileDevice -Mailbox $M.DistinguishedName
 ForEach ($Device in $Devices) {
   $DaysSinceLastSync = $Null; $DaySinceFirstSync = $Null; $SyncStatus = "OK"
   $DeviceStats = Get-ExoMobileDeviceStatistics -Identity $Device.DistinguishedName
   If ($Device.FirstSyncTime) {
      $DaysSinceFirstSync = (New-TimeSpan $Device.FirstSyncTime).Days }
   If (!([string]::IsNullOrWhiteSpace($DeviceStats.LastSuccessSync))) {
      $DaysSinceLastSync = (New-TimeSpan $DeviceStats.LastSuccessSync).Days }
   If ($DaysSinceLastSync -gt 30)  {
      $SyncStatus = ("Warning: {0} days since last sync" -f $DaysSinceLastSync) }
   If ($Null -eq $DaysSinceLastSync) {
      $SyncStatus = "Never synched" 
      $DeviceStatus = "Unknown" 
   } Else {
      $DeviceStatus =  $DeviceStats.Status }
   $ReportLine = [PSCustomObject]@{
     DeviceId            = $Device.DeviceId
     DeviceOS           = $Device.DeviceOS
     Model              = $Device.DeviceModel
     UA                 = $Device.DeviceUserAgent
     User               = $Device.UserDisplayName
     UPN                = $M.UserPrincipalName
     FirstSync          = $Device.FirstSyncTime
     DaysSinceFirstSync = $DaysSinceFirstSync
     LastSync           = $DeviceStats.LastSuccessSync
     DaysSinceLastSync  = $DaysSinceLastSync
     SyncStatus         = $SyncStatus
     Status             = $DeviceStatus
     Policy             = $DeviceStats.DevicePolicyApplied
     State              = $DeviceStats.DeviceAccessState
     LastPolicy         = $DeviceStats.LastPolicyUpdateTime
     DeviceDN           = $Device.DistinguishedName }
   $Report.Add($ReportLine)
 } #End Devices
} #End Mailboxes
[array]$SyncMailboxes = $Report | Sort-Object UPN -Unique | Select-Object UPN
[array]$SyncDevices = $Report | Sort-Object DeviceId -Unique | Select-Object DeviceId
[array]$SyncDevices30 = $Report | Where-Object {$_.DaysSinceLastSync -gt 30} 
$HtmlReport = $Report | Select-Object DeviceId, DeviceOS, Model, UA, User, UPN, FirstSync, DaysSinceFirstSync, LastSync, DaysSinceLastSync | Sort-Object UPN | ConvertTo-Html -Fragment

# Create the HTML report
$Htmltail = "<p>Report created for: " + ($Organization) + "</p><p>" +
             "<p>Number of mailboxes:                          " + $Mbx.count + "</p>" +
             "<p>Number of users synchronzing devices:         " + $SyncMailboxes.count + "</p>" +
             "<p>Number of synchronized devices:               " + $SyncDevices.count + "</p>" +
             "<p>Number of devices not synced in last 30 days: " + $SyncDevices30.count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------" +
             "<p>Microsoft 365 Mailboxes with Synchronized Mobile Devices<b>" + $Version + "</b>"	
$HtmlReport = $HtmlHead + $HtmlReport + $HtmlTail
$HtmlReport | Out-File $HtmlReportFile  -Encoding UTF8

Write-Host ""
Write-Host "All done"
Write-Host ""
Write-Host ("{0} Mailboxes with synchronized devices" -f $SyncMailboxes.count)
Write-Host ("{0} Individual devices found" -f $SyncDevices.count)

$Report | Export-CSV -NoTypeInformation $CSVReportFile
Write-Host ("Output files are available in {0} and {1}" -f $HtmlReportFile, $CSVReportFile)
Start-Sleep -s 4

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
			}
			#endregion  
			#region Option 98) Exit and restart
			98 {
				#      Exit and restart
				Stop-Transcript
				restart-computer -computername localhost -force
			}
			#endregion   
			#region Option 99) Exit
			99 {
				#      Exit
				if (($WasInstalled -eq $false) -and (Get-Module BitsTransfer))
				{
					Write-Host "BitsTransfer: Removing..." -NoNewLine
					Remove-Module BitsTransfer
					Write-Host "`b`b`b`b`b`b`b`b`b`b`bremoved!   " -ForegroundColor Green
				}
				popd
				Write-Host "Exiting..."
				Stop-Transcript
			}
			#endregion   
			default { Write-Host "You haven't selected any of the available options. " }
		}
	}
	while ($Choice -ne 99)
	
}
#region MAIN SCRIPT BODY
######################################################
#               MAIN SCRIPT BODY                     #
######################################################

# Check for Windows 2012 or 2012 R2
if (($ver -match '6.2') -or ($ver -match '6.3'))
{
	$OSCheck = $true
	Code2012
}

# Check for Windows 2016
if ($ver -match '10.0')
{
	$OSCheck = $true
	Code2016
}

# If Windows 2012, 2012 R2 or 2016 are found, exit with error
if ($OSCheck -ne $true)
{
	write-host " "
	write-host "The server is not running Windows 2012, 2012 R2 or 2016.  Exiting the script." -foregroundcolor Red
	write-host " "
	Exit
}
#endregion 



