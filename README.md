# Re-AOL - Retro AOL® Client software
![Python](https://img.shields.io/badge/Python-3.10-green?style=flat-square&logo=appveyor)

This is the AOL3.0 client software from 1995.

Everything you need for connecting to the Re-AOL service is contained in the ZIP file. The AOL3.0 client was installed on a Winblows 10 PC, tested, then packed into the archive. The client software, nor its dependancies, have been altered in anyway from their original state after being installed, with exception of two files:

>./idb/main.idx

>./ccl/TCP.ccl

**main.idx**: contains all the FDO forms, images, and binary data the client needs for interaction with the server. Any new FDO forms will be compiled to the main.idx then uploaded. This will be temporary until the server can properly handle what is called *Download on Demand* (DOD) protocals which will allow us to automagically update the FDO forms whenever a user logs into the service.

**TCP.ccl**: contains the script that tells the client software where to look for the Re-AOL server for connection. The only thing changed in this file was the address to the Re-AOL server. There are other ways to tell the client where to look for a connection, but this was the easiest way that did not require altering firewall settings or a hosts file.


Visit our Discord for more information:

[#faq](https://discord.com/channels/978741297865175081/979473341322788914)

[#general](https://discord.com/channels/978741297865175081/978741297865175085)

*-This readme will be updated with current server information as updates & fixes are rolled out-*



# Re-AOL server information:
An AOL® (America Online) server written in Python (3.10) focusing on the 16/32-bit AOL® 3.0 client versions - for the time being.
> AOL® versions 4+ integration is planned for the future.
>
> **Currently closed source.**

A nostalgic return to the youth of the 90s', and the online community that sparked a copious amount of young adults interests in software development. America Online was one of the foundations that started many on their paths into the computer sciences.
And so I embark on a journey to [re]animate/[re]vive/[re]turn AOL® by writing a server, from scratch, using resources found all over the internet - while also learning Python.

In addition to being an old geek and nostalgic for the good old '90s, my intent for the project is to bring back a medium that sparked kids of all ages interests' in programming. A lot of those people went on to do great things with the skills they acquired during their "haxor" days on AOL.
I wanted to bring back that same medium, and possibly set some young kids on the right path? The latter is more than likely a pipe dream, but you never know.
Whatever the outcome, I hope it can be a place where people like myself will populate chat rooms and bulletin boards with great late-night conversations like the good old days!

## What works:

>- **Account Creation/Registration: 100%**
>>  - Requires a serial number and a password combo to register an account. In the 90s these were on CD covers and floppy disks. They will be provided on GitHub and the Re-AOL website (if one ever gets made?)
>- **Chat Rooms (Roomer): ~95%**
>> - Chat rooms work 100%, however, moderation (title, kick/ban, private/invisible etc) setup GUI is not there - in progress.
>
>- **Instant Messages (Whisper): ~98%**
>> - The ability to block screen names has not yet been implemented.
>> - Some issues with LONG messages causes client crash - keep it short and works fine.
>
>- **View Member Profiles: ~100%**
>- **Create Member Profiles: ~99%**
>> - Reported random client crashes.
>
>- **Locate Members Online: ~100%**
>
>- **Online Clocks: ~99%**
>>  - timezone and localization are not yet implemented in the display time - currently displays the server time
>- **Buddy List™: ~100%**
>>  - As of February 7, 2022 the Buddy List Groups and names settings are complete! You can now add group names for your friends, co-workers, associates, etc which are all customizable. For examble the default Buddy List group name is "Buddies", you can change it to whatever you want (with-in an 11 character limit).
>- **Web Browser: ~100%**
>>  - crashes often, it needs to pass through a proxy that strips out all the HTML that did not exist back in the 90s' and early 2000s', or a new library to replace the one that comes with the AOL® 3.0 installer.

## In progress:

>- Member Services
>- Bulletin Boards
>- E-Mails
>- Find
>- Keyword and Keyword Search
>- Other Content

### What doesn't work:

Most of the online services content is missing as the FDO91 source code for them was all stored on the AOL® servers - these are called "host forms" - so most, if not all of them will have to be written entirely from scratch. To do this I need resources - images of what the Forms looked like when opened in the client.
I have a small chunk of googled images for some content but need more!

***If anyone has knowledge of the FDO91 language, or access to a good set of FDO91 manuals? Let me know please.***

The source code is closed at the moment, primarily because I'm doing this to get back into programming after many years of absence, and to learn Python. This has become my passion and hobby of which I have dedicated an extreme amount of time to. Which for me personally, is extraordinary. I have an annoying bit of A.D.D. that has taken over my life from the beginning. so finishing anything has always been my demon, until this project came into existence that is.

## _Disclaimer_


## Images


## Getting started


### Prerequisites


### Server: New Releases


## Need help?


### Community


### Reporting security issues and security bugs
