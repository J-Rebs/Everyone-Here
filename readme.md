## Purpose of this repository 

To build out a simple script that when run returns information on upcoming calendar appointments in Microsoft Outlook. The tool is intended to look for meetings in the next 24 hours. 

## Dev notes:
- development stopped for now after initial prototype, further refactoring can be done at a later point 

## Bug notes

- if I am not the organizer of the meeting it shows everyone as “not responded to the invite“ when I actually can not know this information

## How to use

If you would like to use the executable, go to the `exectuable` folder in this repository and download the `check_email.exe` file. 

## Who can use this

This script is free to use, but it has only been tested on a Dell Windows 11 machine. It is assumed not to work with Macs or with non-Outlook email clients. Furthermore, extensive testing of edge conditions is not complete. Use this executable at your own risk.

A second version is under development to address two bugs which are:
 - capture 24 hour time frame (current restriction does not always capture full time frame properly)
 - clearer response messages to avoid misleading print out
 
Revised code is under a `version2` branch. See Issues for more detail.

## Basic design 

Class: Check 

    Methods: get_meeting_info 

## Possible Future additions

1. Graphical interface with QT 
2. Add an icon for the executable 
3. Better how to video with louder audio


## Have feedback for me?
Feel free to send me an email at jr4162@columbia.edu or find me on [LinkedIn](https://www.linkedin.com/in/joe-rebagliati-4ab7a488/).

Refer to issues for current bugs being resolved.

## Video Guide

Below is a video guide that walks through how to use the `check_email.exe` file and how to get it onto your computer. As one correction, I note in the video that if it doesn't display a name in the console but just an email, then the reciepient isn't on Outlook or in your Outlook organization. I'm not sure this is actually the case.

Also I know the audio is quiet, so you may need to use headphones. Thanks!

[![Watch the video](https://img.youtube.com/vi/gy4-afE_uHw/0.jpg)](https://www.youtube.com/watch?v=gy4-afE_uHw)
