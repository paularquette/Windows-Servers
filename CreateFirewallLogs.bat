rem Batch Script to Create Firewall Log Files
rem Written By: Paul Arquette
rem Last Modified: Oct 24, 2022
rem Last Modified For: Github

if exist C:\Windows\System32\LogFiles\firewall\pfirewall.log (
  echo file exists
) else (
  netsh advfirewall set allprofiles logging filename %systemroot%\System32\LogFiles\firewall\pfirewall.log
  netsh advfirewall set allprofiles logging maxfilesize 32767
  netsh advfirewall set allprofiles logging droppedconnections enable
  netsh advfirewall set allprofiles logging allowedconnections enable
)
