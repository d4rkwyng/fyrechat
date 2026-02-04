# FyreChat Help

## Required Files

Included:
- BCEL.dll

Needed:
- Required Visual Basic Files (`comdlg32.dll`, `mscomctl.dll`, `msvcp50.dll`, `mswinsck.ocx`, `richtx32.ocx`)
- Required Visual C++ Files (for plugins)

These are included in `assets/required_files.zip` and `assets/vc7_crt.zip`.

## Installation

1. Extract all files from FyreChat.zip into a folder
2. Execute FyreChat.exe
3. Open the Configuration and fill in your information
4. Select your product and its CDKey
5. Browse through the other options and change if desired
6. Click connect or press F1

## Idle Options

| Variable | Description |
|----------|-------------|
| `%ver` | Version info |
| `%uptime` | Computer uptime |
| `%botuptime` | Application uptime |
| `%connected` | Connection uptime |
| `%idle` | Idle uptime |
| `%self` | Your username |
| `%chan` | Current channel |

## Console Commands

Usage: `/command`

| Command | Description |
|---------|-------------|
| `/rejoin` | Joins Void and returns |
| `/reply` *text* | Sends text to last user who whispered you |
| `/disconnect` | Disconnects all open connections |
| `/reconnect` | Initiates connection procedures |
| `/uptime` | Displays idle, connect, app, & computer uptime |
| `/botuptime` | Sends bot uptime to channel |
| `/compuptime` | Sends computer uptime to channel |
| `/connected` | Sends connected uptime to channel |
| `/recorddata` | Displays system info from server |
| `/profile` *user* | Retrieves the profile of user |
| `/version`, `/about` | Displays bot version information |
| `/botnews` | Retrieves news from FyreChat website |
| `/setusername` *user* | Sets the username in configuration |
| `/setpassword` *password* | Sets the password in configuration |
| `/setserver` *server* | Sets the battle.net server in configuration |
| `/setproduct` | Displays available products to set to |
| `/setproduct` *product* | Sets the product in configuration (e.g. `/setproduct STAR`) |
| `/sethome` *channel* | Sets the channel in configuration |
| `/setcharactername` | Sets the Realm character name in configuration |
| `/realmconnect` | Connects to Realm |
| `/getlag` | Gets ping of self |
| `/getlag` *user* | Gets ping of user |
| `/clear` | Clears all buffers (chat, whisper windows, and send textbox) |
| `/clearscreen` | Clears chat screen |
| `/queue` | Queue size |
| `/clearqueue` | Clears the queue of messages |

## Plugin Commands

| Command | Description |
|---------|-------------|
| `/loadplugin` *plugin* | Loads the plugin (e.g. `/loadplugin plugins\aliases.bcp`). Plugins in the plugins folder are automatically loaded. |
| `/unloadplugin` *plugin* | Unloads a loaded plugin |
| `/listplugins` | Lists the plugins that have been loaded |

## Experimental Commands

| Command | Description |
|---------|-------------|
| `/downloadfile` *filename* | Downloads filename (e.g. `icons.bni`, `bnserver.ini`) |
| `/news` | Gets and displays news from battle.net 0x46 |
| `/realmserverlist` | Displays available realms to connect to |
| `/realmgamelist` | Displays available Diablo 2 games to join |
| `/realmjoingame` *gamename* | Attempts to join the Diablo 2 game |
| `/changepassword` *newpassword* | Changes your password (currently only WAR3/W3XP) |

## Thanks (Archival)

- Zonker - Helps with all the bugs
- Raihan - Old web logging system and programming help
- UserLoser - Helped with hashing and more
- Fallen - Programming and beta help
- Exorcizt - Extensive beta testing during FyReBoT development
- Damian - First beta tester
- Thaw - Bug reports
- Levaris - Early FyReBoT beta testing
- Cryptic - Old FyReBoT graphic designer and beta testing
- Deep - FyReBoT beta testing
- idiot - FyReBoT bugs and ideas
- Murder - FyreChat beta and ideas
- Goffy59 - FyreChat beta testing
- Marshall - FyreChat beta testing
- Feanor - Moral Support
- Venox - PHP Script for Web Logger
