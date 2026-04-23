# London Siege
London Siege is a turn-based arcade-style game built in Microsoft Excel VBA.

*Version: v2026.04.23.01*

**DEVELOPER NOTE**: I initially created London Siege in 2018-19 while working an administrative job with a lot of downtime. I "knew" how to code, but was unfamiliar with best practices, OOP, etc., and I wasn't even sure what was going to be possible when I started it. Some years on, I've noticed "I built a game in Excel" to be an occasionally piquant phrase when speaking with other devs, so I wanted it available for observation and engagement.

The codebase is currently in a state of transition as I make periodic improvements, DRY it out, eliminate magic numbers, collapse multi-dimensional arrays into User-Defined Types and Enums, etc. And I'll admit, as I'm refamiliarizing myself with the code and game, I don't quite remember all of the behavior, so forgive (and message) me if something in this guide is inaccurate or missing. However, the game is playable, although yet-to-be-developed features make it ostensibly impossible to win all ten levels. There isn't even an endgame yet.

Alas, enjoy!

## Concept
You are charged with defending London over ten nights of aerial raids. You will meet the enemy waves of fighters and bombers with air defense turrets and your own RAF fleet.

## Setup
London Siege requires Microsoft Excel (compatible with the latest version from Microsoft 365) and can be played directly from the downloadable .xlsm file. Open the file locally. Make sure that macros are enabled and Design Mode is off (only four buttons should be visible at game start).

## Starting the Game
Click the 'Start' button (it should be the only one active) to load the game. You will first be prompted to select locations for three turrets. Click a checkbox to add a turret in that position; it will automatically be replaced with an "A" icon.

Click "Night 1" to begin the game.

## Structures
All structures sit on the bottom two rows of the game grid. The first row of structures is randomly generated at the start of each game. The second contains turrets whose positions are selected by the player.

Structures can get damaged by enemy fire, though they can be repaired between nights (see "Between Nights: Repair Services") if not fully destroyed. Structures have a health of 4, and damage is displayed via background color:

| Health | Background |
| ----------- | ----------- |
| 4 | Grass or Horizon |
| 3 | Yellow |
| 2 | Orange |
| 1 | Red |

Destroyed structures will become rubble.

### Turrets - A
You are given three defense turrets, denoted with the icon "A". Each turret can launch a single shot "." in one of eleven directions per turn. The shot must land exactly on the grid square with an enemy plane or fire to hit/intercept. Gameplay continues as long as one of these turrets (or any RAF) survive.

### Repair Services - \#
There are two Repair Services ("#"). These must be kept intact in order to repair your structures and RAF between rounds (see "Between Rounds: Repair Services"). If only one remains, repair prices will double.

### Ammo Bunkers - $
There are two Ammo Bunkers ("$"), but most of their intended functionality has not been coded yet. However, they are required to be able to purchase rockets for the RAF. If only one remains, rocket prices will double.

### Airfield - ____
The Airfield serves as parking for your RAF fleet, denoted with four "_" icons. It is empty to start, but RAF can be purchased between nights (see "Between Nights: RAF"). At the start of a night, you can only have as many RAF planes as there are available spaces in the Airfield. If, during a night, Airfield spaces are destroyed such that there are more RAF planes than Airfield spaces, the RAF fleet will be limited to the number of available spaces.

### Command - *
The Command, denoted with an asterisk "*", allows you to launch RAF planes prior to the arrival of the next wave. When you have RAF planes, clicking the "Night #" button will prompt a dialog box (see "Gameplay: RAF Pre-Launch"). If the Command is destroyed, this dialog box will no longer appear at the start of a night.

### City Structures - l=l
City Structures ("l=l") have no effects other than providing score bonuses at the end of a night.

## Aircraft

### Enemy Fighters - %
Enemy Fighters ("%") are nimble and quick. They will fire anytime they are headed downward toward London or in the general direction of an RAF plane. Their fire only does a single point of damage, and they only take a single shot to destroy.

### Enemy Bombers - {-^ or ^-}
Enemy Bombers ("{-^" or "^-}") are slower than the Fighters and move more predictably, always either right or left across the screen depending on the icon direction (though they may move up and down as well). They don't fire every turn, but their periodic bombs ("!" with a tail "|") do 2 points of damage to structures directly hit and 1 point of damage to structures on either side. Bombers also take two hits of normal turret or RAF fire, and will become ("{-\~" or "\~-}") when damaged.

### RAF - ><
RAF are purchasable between nights, up to the number of surviving Airfield spaces. RAF move at a speed of 100, 150, or 200 m/s, and fire in the direction of movement. RAF at a speed of 100 m/s can turn up to 90 degrees, but only 45 degrees at a speed of 150 or 200 m/s. RAF take 3 hits to destroy, though a direct hit from a dropped bomb or a collision with an enemy plane will also destroy it (though RAF may adjust to avoid collisions).

RAF can be pre-launched before the night (if the Command is still intact) or left in the Airfield and launched during the night, though RAF will get destroyed when sitting on a piece of the Airfield that gets destroyed.

RAF can be equipped with up to 6 rockets (see "Between Nights: RAF") which can be used in place of normal fire on any turn (check "Fire" and the option to use a rocket instead will appear). Rockets ("+") will move twice as fast across the screen and do enough damage to destroy an Enemy Bomber in one hit.

## Gameplay
The game is divided into ten nights, each with a new wave of Enemy Fighters and Bombers. The game is turn-based, and each turn will first prompt aiming of the turrets, followed by controls for each of the RAF. Once control decisions for all of the existing defenses have been made, there will be 24 frames of animation showing the result, including the movement/attack of enemy aircraft. Play continues until all enemy aircraft have been destroyed, or no turrets/RAF remain.

### RAF Pre-Launch
If you have RAF in the Airfield and your Command ("*") is surviving, an RAF Pre-Launch dialog box will appear after clicking "Night #" to start the next night. Click yes to go to the RAF pre-lanch menu in the righthand panel. Select one of the RAF planes via checkbox. Then click a space in the grid and click "Launch Fighter". A warning will appear if the selected cell is beyond the valid range to launch a fighter.

## Between Nights
When the last Enemy Fighter or Bomber is destroyed, the round immediately ends. Bonuses are awarded based on number of turns used, number of shots taken, number of surviving city structures, and an additional night completion bonus. Depending on what structures remain, shop portals will appear:

### Repair Services \[#\]
As long as at least one of your repair services remains intact, the Repair Services "#" button will appear after a night. Clicking on it takes you to the repair menu. Damaged structures are displayed in the righthand panel. Use the corresponding spin buttons to repair structures, and click "Make Repairs" to confirm. Repairs cost points; prices double when one of the Repair Services buildings is destroyed. You will not be able to spend more points than you have.

### RAF \[><\]
If at least one space of the Airfield remains, the "><" button will appear. This will open the RAF menu in the righthand control panel. The menu provides options to purchase new RAF (via checkboxes Airfield space allows it), repair existing RAF (if repair services intact, using spin buttons above RAF), purchase rockets (if ammo bunkers intact, using spin button by storage), and arm/disarm RAF with rockets (up to six per plane).

### Ammo Bunker \[$\]
"$" button will appear when at least one Ammo Bunker remains. Functionality not coded yet, clicking this button will do nothing.

## Winning The Game
You can't. Yet. But let me know your highest score!
