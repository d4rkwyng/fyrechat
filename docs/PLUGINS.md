# FyreChat Plugins

To manually load a plugin, type `/loadplugin path/pluginname.bcp`
(e.g. `/loadplugin plugins/disabled/oper.dll`)

Or move the plugin into the plugins folder for auto-loading.

Plugins require Visual C++ 7 Runtimes (included in Required Files).

*The following plugins have been tested with FyreChat v2.1.0 and higher.*

## Available Plugins

| Plugin | Author | Description |
|--------|--------|-------------|
| Aliases.bcp | Yoni | Create your own aliases for commands or chat-shortcut commands |
| AlwaysOnTop.bcp | Spht | Toggle the bot being displayed on top of other windows |
| AutoRejoin.bcp | Spht | Forces the bot to return to the channel it was kicked from |
| ChatEncode.bcp | Spht | Adds `/hex`, `/encrypt`, and `/reverse` commands for message encoding |
| ChatFilters.bcp % | UserLoser | Manage user and message filters |
| ChatFocus.bcp | Spht | Open conversation windows with Battle.net users via `/chatfocus` |
| ClanManager.bcp | Spht | Add and remove clan members (Warcraft III only) |
| ClanMOTD.bcp | Spht | View/set clan message of the day via `/motd` and `/setmotd` (Warcraft III only) |
| DisableAwayIdle.bcp # | Spht | Disables the away idle (conflicts with FyreChat's internal /away) |
| DMBotOp.bcp | Spht | Adds DMBotOp 1.1 functionality. Do not use with Oper. |
| EmailManager.bcp | Spht | Change account email or request a new password |
| FavoriteChannels.bcp | UserLoser | Manage and join favorite channels |
| FilterOfflineFriends.bcp % | Yoni | Toggle display of offline/non-mutual friends via `/fof` and `/fmf` |
| FriendsManager.bcp | Sopht | Add and remove friends |
| ImmediateSay.bcp | Spht | Say command that bypasses flood protection (`/say`) |
| KillExtraWork.bcp | Spht | Ignores ExtraWork messages when connecting (0x4A, 0x4C) |
| KillFriends.bcp | Spht | Hides binary friends update messages (0x65-0x69) |
| KillRealmMOTD.bcp | Unknown | Ignores realm messages when logging in with a character |
| NetworkAnalyzer.bcp | — | View bytes received from Battle.net via `/networkstats` |
| oper.dll | Spht | Adds bot capabilities to the chat client. Do not use with DMBotOp. |
| SetFont.bcp | Spht | Toggle chat output window font settings |
| StarcraftGameList.bcp | Skywing | List StarCraft games: `/listany`, `/listmelee`, `/listffa`, `/listladder`, `/listums`, `/listtvb` |
| Trivia.zip | Spht | Trivia game with score tracking via `/trivia on\|off\|auto` |
| UiNotify.bcp % | Unknown | Manage UI alerts for talk/emote phrases |
| WinampControl.bcp | UserLoser | Control Winamp: `/play`, `/stop`, `/pause`, `/next`, `/back`, `/volup`, `/voldown`, `/rewind`, `/forward`, `/closewinamp`, `/songinfo`, `/sendsonginfo` |

**%** — Loads correctly but does not work.
**#** — Loads correctly but causes issues.

## Write Your Own

The BCP SDK is available in `assets/` and includes EmptyPlugin, PluginMenu, BCP Header, TelnetServer API, and BinaryChatCoreLite.
