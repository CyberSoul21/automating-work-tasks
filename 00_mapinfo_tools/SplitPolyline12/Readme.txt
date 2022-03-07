Split Polyline 1.2 for MapInfo

by Nathan Woodrow
22/08/2006

Moves along Polyline to a user specified distance and places a new node,
then uses a function to split the line at the node just added.
Note: Only works for single segment Polylines(not combined)

Contains code taken from the LINEDIR.MBX utility created by Jennifer Bailey of MapInfo Technical Support


Version 1.2
Changlog
	-Fixed bug with no. of sections when selecting a line
	-Fixed problem when switching line direction and spliting

Version 1.1
Changlog
	- works with all projections (Spherical only)
	- improved error handling, negative or invalid values entered etc.
	- uses the measurement unit of the current window when prompting for a distance.
	- automatic temporary linestyle change to show direction of selected line. 
	- added 'Reverse Line' button to dialog for easier line direction changing.
	- no longer requires a map window to be open before launching.
Version 1.0
	- initial release

This is a program to split a polyline at a user specified distance from the start of the line.  

The Split Polyline tool can also be accessed via a tool added to the 'Tools' buttonpad.

Because it measures from the start of selected polyline, the selected line is temporarily changed
to show its direction. A button has been added to the dialog to  reverse the direction of the selected
line if needed.

All your base are belong to us!
