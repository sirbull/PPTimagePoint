#!/bin/bash
# Installerer PPTimagePoint i PowerPoint for Mac.
DIR="$(cd "$(dirname "$0")" && pwd)"
DEST="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$DEST"
cp "$DIR/manifest.xml" "$DEST/PPTimagePoint.xml"
echo ""
echo "PPTimagePoint er installert."
echo "Avslutt PowerPoint helt (Cmd+Q) og start det på nytt."
echo "Gå så til Hjem > Tillegg > Mine tillegg og velg PPTimagePoint."
echo ""
read -n 1 -s -r -p "Trykk en tast for å lukke vinduet."
