# PnIdLineGroupFromToHelper
Example Code - use at own risk

created by gpt5.4 pro (extended) - seems to work as prompted in a small test (no corrections / iterations were needed) - check code before using, use totally at own risk

<img src=FromToTool.png>

prompt used:

Hi, I need a C# based tool for Plant 3D P&ID drawings that does the following:
prompt for a pipeline segment to be selected.
after selection read from that pipeline segment the pipeline group
now loop through all pipeline segments of that group and find connected equipment
if less than 2 equipment found, cancel the command with the message, that at least 2 equipment need to be connected to the line group for this command.
if there are exactly 2 equipment found, suggest a "from" and "to" assignment of the equipment, something like this: "pipeline group <detected line group> goes from <tag of suggested>
from equipment and drawing file name, if eqp not on this drawing> to <tag of the other equipment>, is this correct (yes/no/cancel)?" Please make yes/no/cancel clickable.
If the suggestion is accepted by <enter> then the tag information of the from and to equipment is written into pipe line group attributes with the names: from and: to.
if the suggestion is not accepted ("no"), the from and to information will be flipped and a new suggestion will be made, just as explained above.
"cancel" of course means that the whole script execution will be stopped, also a second "no" will lead to cancellation.
if more than 2 equipment found connected to the line group, please offer the user to select the "from" equipment first, and then the "to" equipment and write the gathered information
info the line group properties as explained above. The "from" and "to" options are given by command line as clickable options, but if the equipment is in the active drawing, then
also selecting the equipment object in the drawing by mouseclick is valid.
