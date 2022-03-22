# get-switch-info
  * Install Remote Server Administration Tools for Windows 10
    * Link: https://www.microsoft.com/en-us/download/details.aspx?id=45520
  * Install powershell version >= 7.0
    *  Link: https://github.com/PowerShell/PowerShell/releases/tag/v7.2.2
  * Install Powershell Module PSDiscoveryProtocol
  ``` Install-Module -Name PSDiscoveryProtocol ```
  * Execute get-switch-info.bat with admin privilege
    * Check the output to see the all the Pcs have right info in the table.
  * Copy the GLPI.ini to root of the c:\
    * Fill in with the appropriate APItoken & AppToken
  * Make sure on the GLPI you have created the switch and networkoutlet.
  * Launch the update_network_info_to_glpi.bat to update the info to GLPI
 


