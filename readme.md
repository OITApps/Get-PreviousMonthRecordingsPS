


# Get-PreviousMonthRecordings Powershell

Many users of our platform have compliance or other requirements that necessitate storing their call recordings offline. This PowerShell script will retrieve all call recordings for the previous month.

## Purpose

Many users of our platform have compliance or other requirements that necessitate storing their call recordings offline. This PowerShell script will retrieve all call recordings for the previous month.

## Installation
Save this file in a folder where you want call recordings to be stored. Edit the configuration block at the top to add your

- FQDN
- Domain
- Client ID
- Client Secret
- PBX user name
- PBX password

If you do not have these details, you can obtain them from support@oit.co. Note the PBX user and password should be dedicated for the application.

## Use

Run this file from the folder and it will do the following:
- Create sub folder with the previous month's YYMM
- Save all recordings from the previous calendar month in the folder
- Create a manifest CSV with all calls and recording file names referenced.
It is recommended to use Task Schedule or a Cron job to run this monthly.

## Support

This script is provided as is, without any warranty or SLA. If you have an issue, contact OIT and we will provide support on a best-effort basis.
