# UFTLogs.vbs
VBScript UFTLogs.vbs is used to assist in setting HPE Unified Functional Testing (UFT), UFT License, UFT API, and other logs

## Purpose

When trouble-shooting issues with UFT, sometimes support will ask to have UFT logs placed in debug mode.  This can be done using HPE procedure documented in   [KM00467327](https://softwaresupport.hpe.com/group/softwaresupport/search-result/-/facetsearch/document/KM00467327).  This requires finding an xml file and making changes for capture and then changing back after capture is complete.


#### Issues addressed

1. Manual changing of the xml can lead to time lost and manual errors.
2. Confusion on what items were changed and needs to be changed back.
3.  TODO More information?

## Overview of Code steps
Runing of this code can be done by clicking on the VBS file.

    Note: If the company's firewall policy does not allow downloading of VBS files,
    copy the  text and place it into Notepad then save the file with .vbs extension.
Don't run this while UFT is running.

#### Following items are addressed with this code.
1. Check to see if UFT is running.  Script will ask you to close UFT and exit.
2. Check to see if UFT remote agent is running.  Script will kill agent if running after notifying.
3. Ask to delete UFT current logs.
4. If backup does not exist, creates a back up of the log.config.xml file as log.config.xml.BAK in the same directory.
5. Asks to set UFT Core, License, Help, HTML Report Logger, Mobile and API in debug logging mode. If user answers "Yes", log mode will be set to Debug. If user answers "No", log mode will remain or be reset to "Error".

### Usage:

usage:   UFTLogs.vbs


### Thank you
Thank you to Michael Deveaux for testing and reviewing code.
