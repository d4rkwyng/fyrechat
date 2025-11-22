FyReBoT Read-Me
< Updated for 2.0.55 >
-------------------------------------------
- Brief EULA (End-User License Agreement) -
-------------------------------------------
- This program can be edited in any means necessary
- This is freeware software and it is to be used on your own risk
- The Author and/or FSI takes no responsibility
- Full distribution of the public version program is allowed

* Please scroll down to Important Information at the bottom of the page for more

------------------
- Required Files -
------------------
ARDenc.dll   - Included
FyReBoT.dll - Included
comdlg32.dll - ?
comdlg32.ocx - ?
mscomctl.dll - http://fyrebot.efyre.com/required
msvcp50.dll  - " "
mswinsck.ocx - " "
richtx32.ocx - " "
Hash Files   - http://fyrebot.efyre.com/hash

---------------------
- Bot Configuration -
---------------------
< Config >
 < Connection >
	Username         - Login Server Name
	Password         - Login Server Password
	Server           - Login Server
	Channel          - Channel to join on Server Connection
	Product          - Blizzard Product you wish to Emulate
	CD-Key           - CD-Key for the Emulated Client
	CD-Key 2         - Additional CD-Key for LOD/TFT
	CD-Key Owner     - Sets the User of the CD-Key
	Ping Spoof
          None           - Sets Normal Ping
          -1ms           - Sets Ping to -1
          0ms            - Sets Ping to 0
        UDP Plug         - Sets UDP Plug (Not compatible with D2DV/D2XP)
	Ping Delay       - Increases Ping with the Set ammount
	Character        - Realm Server Name
	Realm            - Realm Server
	Connect to Realms- Automatically connects on Server Connection
 < Options >
	Join/Leave       - Join and Leave Display Notices
        Diablo 2 Status..- Diablo 2 Status Update Display Notices
	Ping:Flags       - Ping and Flag Display Notices before Chat
        Log Chat         - Logs all events
	Max Length       - Sets Max Ammount of Characters per RTB
	Connect on Start - Automatically Connects On Application Start
	Minimize to Tray - Application goes to tray when minimized
        Confirm Exit     - Will confirm shutdown if connected
	Auto-Reconnect   - Re-Connects on Disconnect
	Self Flood Prot  - Flood Protection to avoid Flooding Server
        Join Product     - Joins a Product-Specific channel before Home
        Normal
	  Normal         - Displays message depending on Idle Interval
	  Idle Interval  - The idle's interval in seconds
        More
	  Hmm            - Sends Hmm on 35 talks including emote event
	  Away Uptime    - Set your away as idle-uptime	
 < Appearance >
        Self-Explanitory For Now
 < Hashing >
 Hash Information:
  Make sure that you have all the hash files and in their proper locations.
  You can edit the hash location in the Configuration section. Make sure the the
  required files are in the chosen folder.
  Suggestion: Create a folder Hash in your FyReBoT directory and make sub folders
  for each Client based on the statstring (STAR - Starcraft, SEXP - Brood War, D2XP
  - Lord of Destruction, W2BN - WarCraft II BNE)
  Need Hash files: http://fyrebot.efyre.com/hash
  If you don't wish to use hash files then FNLS is available and is recommended for
  beginners.

 < Filters Manager >
    < Add/Delete User >
        Add              - Adds user to Filters (Wildcarding allowed)
                           (Ex: *Flood* will ban 2Flood22003.. etc..)
        Delete           - Removes user from Fitlers
        Rejoin Amount....- Automatically filters user if they Join set
                            number of times
< End Config >

---------------------------
- Commands (Console-Side) -
---------------------------
< Commands >
 < General >
	/uptime          - Displays idle, connect, app, & computer uptime
	/botaway         - Gets how long bot has been away
	/rejoin          - Joins Void and Returns
	/rejoin2         - Other means of rejoining
	/version         - Displays bot version information
	/about           - " "
	/clear           - Clears Everything
	/recorddata      - Displays personal info
	/e <text>        - Encrypts in X-Base64 (Thanks to Zonker[RC])
	/h <text>        - Encrypts in Hex format
	/n <text>        - Reverses text
	/r <text>        - Sends text to last user who whispered you
	/clearqueue      - Clears the queue of messages
	/queue           - Queue size
	/banned          - Number of Banned Users
	/bannedusers     - List of all Banned Users
	/filteredusers   - List of all Filtered Users
	/count           - List of all Channel Users
	/wa <command>    - Winamp Commands
	    back         - Plays prev track
	    play         - Play or Re-Plays
	    pause        - Pauses or Unpauses Winamp
	    stop         - Stops the current track
	    next         - Plays next track
	    status       - Shows current played track
	    first        - Plays first track in playlist
            last         - Plays last track in playlist
	    music        - Outputs current track to bnet channel
 *wildcard exceptions
  (Ex: /m FyRe* Hey) If a user in the channel has the name FyRe.. then he will be messaged.
	/m               - Allows Wildcarding and Diablo 2 Support
	/w               - " "
	/msg             - " "
	/whisper         - " "
	/tell            - " "
	/ignore          - " "
	/squelch         - " "
	/unignore        - " "
	/unsquelch       - " "
	/whois           - " "
	/whereis         - " "
	/profile <user>  - User's Profile information
	/getinfo         - Your User Info
	/getinfo <user>  - Other User's Info
	/getlag          - Your Ping
	/getlag <user>   - Other user's Ping
< End Commands >

----------------
- Idle Options -
----------------
 %ver       - Version Info
 %uptime    - Computer Uptime
 %botuptime - Application Uptime
 %connected - Connection Uptime
 %idle      - Idle Uptime
 %self      - Your Username
 %chan      - Current Channel
 %music     - Current Winamp Track

-------------------------
- Important Information -
-------------------------
When using this application certain computer information is logged to a server
for recording/testing purposes. Make sure this program was downloaded from
FSI's web-site or a trusted partner.  FSI (http://www.efyre.com)
Trusted:
 PSGaming (http://www.psgaming.net)
 Madz (http://www.madz.tk)
 NetBlues (http://www.netblues.org)
 ValidX (http://www.validx.net)

----------
- Thanks -
----------
Thank you for trying out my application!
Please report all bugs to: fyre@efyre.com

For more information visit my web-site: http://www.efyre.com

Copyright ©2003-2004 Fyre Systems, Inc. All Rights Reserved.