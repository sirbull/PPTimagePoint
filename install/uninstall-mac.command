#!/bin/bash
# Fjerner PPTimagePoint fra PowerPoint for Mac.
rm -f "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/PPTimagePoint.xml"
echo ""
echo "PPTimagePoint er fjernet. Start PowerPoint på nytt."
echo ""
read -n 1 -s -r -p "Trykk en tast for å lukke vinduet."
