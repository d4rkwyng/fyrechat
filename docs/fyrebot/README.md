# FyReBoT Read-Me

*Updated for 2.0.55*

## EULA

- This program can be edited in any means necessary
- This is freeware software and it is to be used at your own risk
- Mindtricks takes no responsibility
- Full distribution of the public version program is allowed

## Required Files

| File | Status |
|------|--------|
| ARDenc.dll | Included |
| FyReBoT.dll | Included |
| comdlg32.dll | Required |
| comdlg32.ocx | Required |
| mscomctl.dll | [Download](https://mindtricks.io) |
| msvcp50.dll | [Download](https://mindtricks.io) |
| mswinsck.ocx | [Download](https://mindtricks.io) |
| richtx32.ocx | [Download](https://mindtricks.io) |
| Hash Files | [Download](https://fyrechat.openfyre.net/hash/) |

## Bot Configuration

### Connection

| Setting | Description |
|---------|-------------|
| Username | Login Server Name |
| Password | Login Server Password |
| Server | Login Server |
| Channel | Channel to join on Server Connection |
| Product | Blizzard Product you wish to emulate |
| CD-Key | CD-Key for the emulated client |
| CD-Key 2 | Additional CD-Key for LOD/TFT |
| CD-Key Owner | Sets the User of the CD-Key |
| Ping Spoof: None | Sets Normal Ping |
| Ping Spoof: -1ms | Sets Ping to -1 |
| Ping Spoof: 0ms | Sets Ping to 0 |
| Ping Spoof: UDP Plug | Sets UDP Plug (Not compatible with D2DV/D2XP) |
| Ping Delay | Increases Ping with the set amount |
| Character | Realm Server Name |
| Realm | Realm Server |
| Connect to Realms | Automatically connects on Server Connection |

### Options

| Setting | Description |
|---------|-------------|
| Join/Leave | Join and Leave Display Notices |
| Diablo 2 Status | Diablo 2 Status Update Display Notices |
| Ping:Flags | Ping and Flag Display Notices before Chat |
| Log Chat | Logs all events |
| Max Length | Sets Max Amount of Characters per RTB |
| Connect on Start | Automatically Connects On Application Start |
| Minimize to Tray | Application goes to tray when minimized |
| Confirm Exit | Will confirm shutdown if connected |
| Auto-Reconnect | Re-Connects on Disconnect |
| Self Flood Prot | Flood Protection to avoid Flooding Server |
| Join Product | Joins a Product-Specific channel before Home |
| Normal Idle | Displays message depending on Idle Interval |
| Idle Interval | The idle's interval in seconds |
| Hmm | Sends Hmm on 35 talks including emote event |
| Away Uptime | Set your away as idle-uptime |

### Appearance

Self-explanatory.

### Hashing

Make sure that you have all the hash files and in their proper locations. You can edit the hash location in the Configuration section. Make sure the required files are in the chosen folder.

**Suggestion:** Create a folder `Hash` in your FyReBoT directory and make sub folders for each client based on the statstring (STAR - StarCraft, SEXP - Brood War, D2XP - Lord of Destruction, W2BN - WarCraft II BNE).

If you don't wish to use hash files then FNLS is available and is recommended for beginners.

### Filters Manager

| Action | Description |
|--------|-------------|
| Add | Adds user to Filters (wildcarding allowed, e.g. `*Flood*` will ban 2Flood22003, etc.) |
| Delete | Removes user from Filters |
| Rejoin Amount | Automatically filters user if they Join set number of times |

## Commands (Console-Side)

### General

| Command | Description |
|---------|-------------|
| `/uptime` | Displays idle, connect, app, & computer uptime |
| `/botaway` | Gets how long bot has been away |
| `/rejoin` | Joins Void and Returns |
| `/rejoin2` | Other means of rejoining |
| `/version`, `/about` | Displays bot version information |
| `/clear` | Clears Everything |
| `/recorddata` | Displays personal info |
| `/e <text>` | Encrypts in X-Base64 |
| `/h <text>` | Encrypts in Hex format |
| `/n <text>` | Reverses text |
| `/r <text>` | Sends text to last user who whispered you |
| `/clearqueue` | Clears the queue of messages |
| `/queue` | Queue size |
| `/banned` | Number of Banned Users |
| `/bannedusers` | List of all Banned Users |
| `/filteredusers` | List of all Filtered Users |
| `/count` | List of all Channel Users |
| `/wa <command>` | Winamp Commands (back, play, pause, stop, next, status, first, last, music) |

**Wildcard exceptions:** `/m FyRe* Hey` — if a user in the channel has the name FyRe, they will be messaged. Wildcarding supported for: `/m`, `/w`, `/msg`, `/whisper`, `/tell`, `/ignore`, `/squelch`, `/unignore`, `/unsquelch`, `/whois`, `/profile`, `/getinfo`, `/getlag`

## Idle Options

| Variable | Description |
|----------|-------------|
| `%ver` | Version Info |
| `%uptime` | Computer Uptime |
| `%botuptime` | Application Uptime |
| `%connected` | Connection Uptime |
| `%idle` | Idle Uptime |
| `%self` | Your Username |
| `%chan` | Current Channel |
| `%music` | Current Winamp Track |

## Important Information

When using this application certain computer information is logged to a server for recording/testing purposes. Make sure this program was downloaded from [Mindtricks](https://mindtricks.io) or a trusted partner.

## Contact

For bugs or suggestions, open an issue on GitHub or contact @d4rkwyng.

For more information visit [mindtricks.io](https://mindtricks.io).
