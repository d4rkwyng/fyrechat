FyreChat is a Binary Chat application developed from FyReBoT source. The application provides an easy to use interface to connect to battle.net emulating a Blizzard client.

#### Features 
- Icon Loading
- Spoof Ping to use -1ms.jpg, 0ms.jpg or Plug
- Custom Idle Messages and Presets
- Create & Join StarCraft games
- Manage Diablo 2 characters
- Create Diablo 2 games (in development)
- System Tray Icon
- Flood Protection
- Logging & Debugging
- Customize fonts and colors
- Plugins
  - Live Channel List
  - Display Channel Source By venox (Kevin Stacey)
  - Archived Channel List

#### Version History

- **2.1 build 9 (23 August 2012)**
    - Added BNCS (C/S 0x55, 0x56) and BNLS (C/S 0x05, 0x06) packet handling for password changes.
    - Added command **/changepassword** to change passwords on NLSv2 (WAR3/W3XP)
    - Temporarily added boolean to check if web channel list had been changed. Removed.
    - Added check for if BNCS 0x0B channel list is empty, if so: resends 0x0B.
    - Fixed failed BNLS 0x11 on asia.battle.net servers. Needed to be "00" for IPs with "0"
    - Updated Create StarCraft Games to closer match client and added WarCraft II creation as well.
    - Fixed browsing for maps in StarCraft/Warcraft game creation window.
    - Fixed bug with minimizing to system tray.
    - **Released to public (08-24-12)**
- **2.1 build 8 (21 August 2012)**
    - Added command **/botnews**. Retrieves news from FyreChat website.
    - Fixed bug with client not reconnecting after creating a new account.
    - Added account creation under WAR/W3XP.
    - Fixed some typos and colors of text.
    - Added new option to hide displaying of Server Types in chat (BNCS, BNLS, etc...).
    - Cleaned up BNCS code.
    - Restructured BNLS code.
    - Temporarily added BNLS (0x0A) to confirm BNCS (S 0x54) proof. Removed. BNLSWarden doesn't support?
    - Added BNLS (0x11) verification of server IP and signature from BNCS (S 0x50) - only WAR3/W3XP