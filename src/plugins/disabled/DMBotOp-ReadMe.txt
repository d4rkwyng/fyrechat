*********************************************************************
* DMBot Channel Operator 1.1 by DarkMinion (as known on battle.net) *
*********************************************************************

---------------------------------------------------------------------
DMBotOp.bcp is a revision 2 BinaryChat plugin which was compiled by Spht using the source code to DMBotOp 1.1 provided by DarkMinion.
This is the original DMBotOp 1.1 "readme" which was written by DarkMinion.  Some things don't apply to the plugin, such as Battle.net configuration instructions and word game.
---------------------------------------------------------------------


Alright kids, listen up.

This channel op is VERY simple to use.  If the bot doesn't find a config file or database file when it loads up, it will create these files with default information. To set up the bot, simply add your name to the database file with the flags you want (for master, you can just add N and then add whatever you want remotely when the bot is connected), then set up the config file using the configuration dialog.  To access the dialog, simple click the Options/Configure... menu item. Then you will be ready to go!

Put your hash files in a folder named Starcraft in your bot directory.

The bot uses a flag-based system to allow users to control it.

The flags are as follows:

A: Limited userlist access (Can change attributes B, F, S, or L of a given user)
B: Autobanned (can be overridden with S)
C: Ability to add/delete/use custom commands
D: Any user with this flag is auto designated and can also access the designate and designated commands
F: Friend (safe from channel protection)
I: Internal function access (reload, reconnect, quit, uptime)
L: Access to find command
N: Full userlist access (Can change any attribute of a given user)
O: Access to operator commands (ban, kick, unban)
S: Safelisted (safe from ban/kick commands, can be overridden with Z)
T: Access to say command

The commands are as follows (along with a short description and needed flags):

Key: <> - indicates the parameter is required
     [] - indicates the parameter is optional

version
 - Displays outloud the bot's version and name

say <text>
 - Repeats the specified text
 - Required flag(s): T

ban <user> [message]
 - Bans the specified user with the optional message
 - Required flag(s): O
 - Supports wildcards (i.e. ban *darkminion*)

kick <user> [message]
 - Kicks the specified user with the optional message
 - Required flag(s): O
 - Supports wildcards (i.e. kick *darkminion*)

unban <user>
 - Unbans the specified user
 - Required flag(s): O

findattr <flag>
 - Returns a list of users with the specified flag
 - Required flag(s): A, L, or N

find <user>
 - Returns the flags for the specified user
 - Required flags(s): A, L, or N

remove <user>
 - Removes a user from the list regardless of flags
 - Required flags(s): N

chattr <user> <+/-flags>
 - Changes the specified user's attributes
   - Examples: chattr testuser B (sets testuser's flags to only B)
               chattr testuser +B (adds flag B to testuser's flags)
               chattr testuser -B (removes flag B from testuser's flags)
 - Required flag(s): A or N
 - Users with only the A flag have limited access to this command
 - The userlist supports wildcards (i.e. chattr *darkminion* NOS)

join <channel>
 - Forces bot to join the specified channel
 - Required flag(s): J

rejoin
 - Forces bot to rejoin the current channel
 - Required flag(s): J

designate <user>
 - Designates the specified user
 - Required flag(s): D

reload
 - Reloads the bot's config file and database remotely
 - Required flag(s): I

protect <on | off>
 - Toggles channel protection
 - Required flag(s): O

designated
 - Displays outloud the currently designated user
 - Required flag(s): D

reconnect
 - Forces bot to reconnect
 - Required flag(s): I

quit
 - Disconnects the bot and closes the program
 - Required flag(s): I

uptime
 - Displays outloud how long the bot has been connected to battle.net
 - Required flag(s): A, I, or N

sysinfo
 - Displays various system information
 - Required flag(s): A, I, or N

set <internal variable> <value>
 - Sets the given variable to the given value
 - Available variables are: server, homechan, idle, trigger
 - Required flag(s): I or N

addcmd <command name> <command syntax>
 - Adds the given command with the given syntax
 - Use is explained below
 - Required flag(s): C

delcmd <command name>
 - Removes the given custom command
 - Required flag(s): C

clearqueue
 - Clears the bot's send queue
 - Required flag(s): I or N

wordgame <on|off>
 - Turns the built-in word game on or off
 - Required flag(s): I

.a <wordgame answer>
 - Gives your answer for the word game
 - Required flag(s): None

Custom commands are pretty simple to add and use.  It uses a simple argument list in which %1 = first argument, %2 = second argument, and so on. Sample:

<DarkMinion[vL]> .addcmd poke /me pokes %1
<DarkMinion> Command 'poke' added.
<DarkMinion[vL]> .poke dm
* DarkMinion pokes dm *

The built in word game is simple to play.  You simply turn it on and start guessing letters.  Simply say the letter you want to guess.  If you think you know the answer prematurely, use the a command.

That's about it folks, the bot is pretty straightforward, good luck.