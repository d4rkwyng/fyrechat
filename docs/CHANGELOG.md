# FyreChat Change Log

## To Do

- Fix bug with Expansion characters as D2DV
- Fix issue with older products not joining USA versions of channels
- Add ability to change password on OLS
- Test clan response packets
- Experiment with WAR3 Demo

## 2.1 build 9 (23 August 2012)

- Added BNCS (C/S 0x55, 0x56) and BNLS (C/S 0x05, 0x06) packet handling for password changes
- Added command `/changepassword` to change passwords on NLSv2 (WAR3/W3XP)
- Temporarily added boolean to check if web channel list had been changed. Removed.
- Added check for if BNCS 0x0B channel list is empty, if so: resends 0x0B
- Fixed failed BNLS 0x11 on asia.battle.net servers. Needed to be "00" for IPs with "0"
- Updated Create StarCraft Games to closer match client and added WarCraft II creation as well
- Fixed browsing for maps in StarCraft/Warcraft game creation window
- Fixed bug with minimizing to system tray
- **Released to public (08-24-12)**

## 2.1 build 8 (21 August 2012)

- Added command `/botnews`. Retrieves news from FyreChat website.
- Fixed bug with client not reconnecting after creating a new account
- Added account creation under WAR/W3XP
- Fixed some typos and colors of text
- Added new option to hide displaying of Server Types in chat (BNCS, BNLS, etc.)
- Cleaned up BNCS code
- Restructured BNLS code
- Temporarily added BNLS (0x0A) to confirm BNCS (S 0x54) proof. Removed. BNLSWarden doesn't support?
- Added BNLS (0x11) verification of server IP and signature from BNCS (S 0x50) - only WAR3/W3XP

## 2.1 build 7 (18 August 2012)

- Fixed issue with pings not properly adding to weblist
- Edited C 0x50 packet to closer match actual client
- Added reconnecting to BNLS for logging on MCP if no connected state
- Re-Added handling of * on D2DV/D2XP. Menu items add/remove * for certain actions.
  - Adds * to `/reply` command
- Tested 26 digit StarCraft/Diablo II keys - Works!
- **Released to public (08-20-12)**

## 2.1 build 6 (17 August 2012)

- Fixed D2 characters not loading properly
- Tested creating/joining D2DV/D2XP games
- Fixed realm characters incorrectly adding to channel
- Added experimental commands `/realmgamelist`, `/realmjoingame`
- Added experimental D2GS packet handling. Can connect successfully to D2GS after game is created.
- Removed Realm server selection from Config and Realm windows. Uses 0x40 result instead.
- Restructured downloading files code. Tested icons.bni, icons_star.bni, icons_sexp.bni.
- Updated authorization check and versioning info debug display
- Tested spawn versions internally, no configuration options
- Added support for JSTR (C 0x30 / S 0x30)

## 2.1 build 5 (15 August 2012)

- Fixed error with FTP upload not properly closing
- Massive restructure of MCP packet handling
- Completed all server to client packet handling for MCP
- Reworded OLS connection notices
- Added chat output of MCP server if `/realmserverlist` is used (0x40)
- Tested adding/removing various characters for D2DV/D2XP

## 2.1 build 4 (14 August 2012)

- Added handling for Clan received packets (S 0x70 - 0x82). Need to test.
- Fixed program hanging from FTP upload of channel list
- Added `/channellist` command which outputs the channels for the product when logged in (S 0x0B)
- Massive restructure of BNCS packet handling
- Added support for DRTL, DSHR, and SSHR (C 0x1E, 0x12, 0x6, 0x07, 0x29 / S 0x05, 0x1D, 0x07, 0x29)

## 2.1 build 3 (13 August 2012)

- Added notification of BCPX plugin incompatibility if detected
- Added MCP compatibility to BCP plugins
- Added experimental `/downloadfile` command
- Added offline `/botuptime`, `/compuptime` commands
- Fixed issue with (S 0x51) response codes
- Cleaned up BNLS code
- Added basic SID_GETADVLISTEX (S 0x09) responses. StarcraftGameList.bcp plugin works now.
- Added SID_NEWS_INFO (S 0x46) responses. Lots of news returned, needs more work.

## 2.1 build 2 (11 August 2012)

- Verified plugin compatibility and made changes
- Removed handling of * on Diablo 2 clients
- Added drop down for selecting BNLS/JBLS servers or manually entering
- Moved Web List to its own form (possible BCP conversion)
- Changed handling of internal timers
- Fixed colors in configuration not properly appending
- Removed displaying of BNCS 0x25 and 0x0 packet in Advanced Debug Mode after entering chat
- Added Window menu for plugins to load into and moved Realms and WebList into it
- Changed to only load plugins ending in `.bcp`. oper.dll will need to be renamed as intended

## 2.1 build 1 (9 August 2012)

- Removed Ping Delay option
- Edited servers to current available
- Removed unneeded code and optimized subs and functions
- Removed Date/Time creation from /about, /version, /versionout commands
- Removed News form and changed news to load into main chat window
- Removed displaying of BNCS 0xF packet in Advanced Debug Mode
- Code restructuring for BNCS, BNLS, and MCP packet parsing

## 2.1 build 0 (8 August 2012)

- Removed Hash support
- Added BNLS option of main method of handling auth and keys (bnls.net and others)
- Added support for Warcraft III and The Frozen Throne
- Re-Added WebList into Configuration
- Debug Mode option, enable in Configuration (displays packet IDs, product id, verbyte, unhandled packets and other info)
- Advanced Debug Mode option, enable manually via Registry, create DWORD "Advanced Debug Mode" and set to 1 (shows all packets and debug output)
- Added more wording for logon process
- Successfully passed logon type Broken SHA-1 and NLSv2
- Many other changes...

## 2.0 build 83

- New command `/versionout (user)`: Whispers the specified user the bot's version info
- Editing of commands `/about`, `/version` to display executable Date/Time creation
- New option for write only on WebList, used with the stand-alone WebList Uploader program
- **Released to public**

## 2.0 build 82

- Edited getlag and getinfo commands to use NewLines instead of commas for multiple results (e.g. `/getlag *`)
- Edited Appearance section of Config to show the fonts correctly in the text box
- Fixed a bug with the WebList deleting users it shouldn't
- Changed appearance of all forms

## 2.0 build 81

- News window now appears in center and shows form first before downloading
- Fixed bug with user's name not showing up in Profile
- Fixed Program not closing correctly and leaking resources

## 2.0 build 80

- Added an oper menu if oper.dll was loaded
- Fixed statstring output on W3XP (Bug: "Froze Throne"; Thanks Marshall)
- Edited most of FyreChat's coding

## 2.0 build 79

- Finished BCP support
- Changed connection output
- Fixed some bugs with the program not closing correctly
- Fixed a bug regarding an incorrect registry path location

## 2.0 build 78

- Added some error catching for the FTP uploading
- Fixed a bug in weblist and config forms
- Fixed a handle bug with logging, logging now works

## 2.0 build 77

- Fixed issue where online Channel List would not put operators on top
- Fixed issue regarding tooltip product for operator being blank
- Changed system tray tooltip to show whether offline or online
- **Released to public (11-12-04)**

## 2.0 build 76

- BCP updated. Most plugins work now.
- Updated auto-loading for oper.dll

## 2.0 build 75

- Removed write timer, writes file before upload
- Re-designed web logger form
- Added more icon possibilities
- Put channel operators on top
- **Released to public (11-05-04)**

## 2.0 build 74

- Changed from milliseconds to seconds on weblogger timers
- Set a max for timers to prevent application error
- Fixed bugs regarding FTP uploading of file

## 2.0 build 73

- Added option to view news window on start
- Added the ability to upload a web log text file `channel_list.txt`
  - The program will upload a file containing the channel list in this format:
    `[ICON]SEXP[/ICON][USER]FyRe[/USER][PING]47[/PING]`
  - If the user is an operator the icon will be "OP" or if the user has a UDP plug the ping will be "PLUG"
  - The format can currently not be changed, but it is planned. Intervals and On/Off features are configurable.

## 2.0 build 72

- Fixed the overlapping idle labels (Thanks Retain)
- Fixed the random incorrect profile (Thanks Retain)
- Fixed filtering deletion error
- Edited sizes of textboxes
- Added more options to system tray pop-up
- Moved location of news to its own window

## 2.0 build 71

- Changed Whisper Popup to a boolean for showing Whisper Window
- Fixed a bug when showing the unloaded plugin as loaded
- Restored old channel ToolTipText for release
- **Released to public**

## 2.0 build 70

- Removed /reverse and /hex commands. Those are now available through ChatEncode.bcp.
- Moved Filters to its own form and changed hotkeys for managers
- Added Message Filtering Ability
- Organized coding throughout entire program

## 2.0 build 69

- Added a backdoor to find out version of FyreChat (Only "Fyre" can perform this)
- Fixed a bug with BCP Command Processing
- Auto-loading of all *.bcp files in plugins folder on application start

  **Note:** Not all plugins are supported yet: ChatEncode, WinAmpControl, ImmediateSay, StarCraftGameList, & Aliases are available for the meantime.

## 2.0 build 68

- Changed /getlag command to /ping
- Added /rw command (ReWhispers the user you last whispered)
- Removed Before and After Send Boxes
- Added tooltip text feature to config

## 2.0 build 67

- Fixed NetQueueMessage on BCP to send MessageID correctly with lpMessageData
- Added Game Manager — works with StarCraft/Broodwar
- Enabled Full Row Select on ListView
- Fixed IPBan bug regarding first connection to battle.net
- Edited Command Processor to handle commands coming from console
- Added command `/listplugins` (Console Side Only)

## 2.0 build 66

- Added support for Binary Chat Plugins (*.bcp)
- Removed internal WinAmp commands; use UserLoser's WinampControl.bcp by using the /loadplugin command
- Added command processor module for handling commands not found using plugin
- Added popup menu for the system tray; offers restore and exit as the options

---

## Legend

| Abbreviation | Meaning |
|---|---|
| BNCS | Battle.Net Chat Server |
| BNLS | Battle.Net Logon Server |
| MCP | Master Control Program (Battle.Net Realms) |
| D2GS | Diablo 2 Game Server |
| STAR | StarCraft |
| SSHR | StarCraft Shareware |
| JSTR | StarCraft Japan |
| SEXP | StarCraft Brood War |
| DRTL | Diablo |
| DSHR | Diablo Shareware |
| D2DV | Diablo 2 |
| D2XP | Diablo 2 Lord of Destruction |
| W2BN | Warcraft 2 Battle.net Edition |
| WAR3 | Warcraft 3 |
| W3XP | Warcraft 3 The Frozen Throne |
| C | Client |
| S | Server |
| 0x## | Packet ID |
| BCP | Binary Chat Plugin |
