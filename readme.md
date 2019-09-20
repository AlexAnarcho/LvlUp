# LvlUp
current version: 0.6

## About
This program tracks the time invested into certain skills.  You can specify any skill and enter progress made on a
 daily basis. There is a simple Level algorithm to determine a level and give some gameification to your progress
. The game works with the Pomodoro technique and counts 2 Pomodoros (50 Minutes) as one hour. While this is a
 simplification in my experience the work done with the Pomodoro technique is more focused and productive. Therefore
 the heuristic has been made.

## Level Algorithm
The level algorithm is very simple. To reach a given level you have to complete n hours, where n is the number of the
 next level squared.

_Examples:_

1 hour to reach level 1 (1 ** 2 = 1);

4 hour to reach level 1 (2 ** 2 = 4);

9 hour to reach level 1 (3 ** 2 = 9);

and so on.

After reaching a level the number of hours (called experience, or EX for short) is reset. Meaning the 1 hour 
for level 1 does not count into the 4 hours of level 2. To reach level 2 you have to complete 4 whole hours more.

## Keeping Data
In the background there is a xlsx sheet created with openpyxl that lists the entries with datetime objects and number
 of hours. Also kept in the sheet are the current level, the next level, the hours required to reach the next level
 , the hours already completed in the current level and the total number of hours.

## Required 3rd party Python modules
* openpyxl
* inquirer
* matplotlib

---
## Patch notes
* version 0.6: added support for matplotlib and 2 simple graphing functions:
    * Hours invested per day
    * Progress over time

* version 0.5: added support for multiple skills to levelup

## Upcoming features
* When starting for the first time, create a new worksheet with openpyxl and let the user specify the path.
    * Give a default path for the user as well.
    * Let the user specify hours already invested in a certain skill
* Don't write a new daytime object if the entered amount is 0
* Implement further graphs to visualize characteristics

---
## Copyrights and contact info
There are no copyrights for this project. Do with the code whatever you please. The creator believes in self-ownership
, just property rights, the resulting freedom of speech and the non-existence of intellectual property. Open Source
 is the way to.
 
If you like the project and/or have ideas for further features, you can contact me at alexanarcho@protonmail.com