# Blasts from the past

This is a collection of scripts I have created years ago and which might still be useful for some. They can be divided into several groups:

## Active Directory

These scripts were used in a project with a quickly growing number of members. New members' data was gathered in Excel files (.xls) and converted into active directory accounts using two of the scripts given below.

- accttest.vbs (VBScript): detects if an active directory account is enabled or disabled
- create_ldapuser.vbs (VBScript): creates a new active directory user account
- excel2ldap.vbs (VBScript): transform excel data into a text file with ldap-like syntax. To be processed further with ldap2user.vbs. Includes a manpage (German language)
- filteraccount.vbs (VBScript): filters accounts by a given attribute
- findaccount.vbs (VBScript): looks up for a specific account
- ldap2user.vbs (VBScript): transforms data produced by excel2ldap.vbs into active directory accounts. Includes a manpage (German language)
- readexcel.vbs (VBScript): reads and displays data from an .xls file. 

## MCMS 2002

I really hope you don't have to work with Microsoft's Content Management System 2002 any longer! This complete CMS fail was discontinued by Microsoft sometime around 2006. These scripts were helpful, though.

- CMSResources-support-job.vbs (VBScript): a support script sending status info to email recipients
cmsexport.vbs (VBScript): exports mcms channels to an .sdo file. Includes a manpage (German language)
- cmshelper.vbs (VBScript): displays information for objects in a given channel. Includes a manpage (German language)
- cmsimport.vbs (VBScript): imports an .sdo file into a mcms channel structure. Includes a manpage (German language)
- resourcewatcher.vbs (VBScript): lists resources of all MCMS resource galleries exceeding a given threshold
- searchresource.vbs (VBScript): looks for MCMS objects by a given GUID, URL, or Channel Path

## Miscellaneous

The "gem" section of this repository, stuff I'm still using today.

- rreboot.cmd: remote shutdown of servers. Gets them into maintenance mode via mmtool call, too.
- TidyUp.vbs (VBScript): for deleting data (temporary files, logfiles etc.). Lots of parameters, see TidyUp.cmd for a sample call.

## Service Control

A useful tool for operating windows services.

- control-service.cmd: for starting, stopping, restarting windows services. Kills processes if stopping doesn't succeed. See start-services.cmd and stop-services.cmd for sample calls.
- restart_component.vbs (VBScript): restarts windows com+ components

## WebSphere 6

Some scripts for operating IBM WebSphere processes, listeners, and more.

- AppServerControl.vbs (VBScript): safely shutdown given WebSphere server by stopping all WAS Application servers on that machine
WasWapTools.txt*

- controlAppServer.jacl (Tcl): start / stop WAS application server, tcl style
- controlAppServer.py (Python): start / stop WAS application server, Python style. See controlAppServer.cmd for a sample call.
- controlListenerPort.py (Python): start / stop a WAS listener port. See controlListenerPort.cmd for a sample call. See readme.txt for sample calls of controlAppServer and controlListenerPort (German language)
- getNodeInfo.jacl (Tcl): print WAS node info
- listRunningServers.vbs (VBScript): enlist application servers with status "running" in a given websphere cell. This is useful if you have to take a "snapshot" about running applications in a cell. running-appservers-T.txt is a sample output file for a given cell. testrun.cmd is a sample call.
- pass.vbs (VBScript): simple tool for "encrypting" passwords that you use in your batches. pwd.ini is a .ini file that can be used for such batches.
- setenv.vbs (VBScript): set / delete passwords, companion of pass.vbs
- shutdownNode.vbs (VBScript): safely shutdown given WebSphere server by stopping all WAS Application servers and node agents on that machine

