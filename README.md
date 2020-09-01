Here's the code I wrote to try to port data from Salesforce SQL files into Excel and back. I suggest anyone adding on to this makes their own repo rather than add to this because I don't really know how to use Git or GitHub beyond uploading files.
Known issues:
-converts dates (both relative and static) into datestrings when loading into Excel and doesn't convert them back
-creates a metadata tab in Excel that users can modify by accident

Uses xlwings (https://www.xlwings.org) and sqlite to load and transfer data, then assembles .sql files as strings.
