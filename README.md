# LabelPrinter
Python console based program to print Warehouse labels in bulk quantities.

First the console prompts the user to input the desired dimensions of the paper, 3 formats are allowed, Once chosen...
Console then prompts user to input the location by asking Bay, Row and amount of Tiers (BAROTI System used in maritime sector to define locations within a vessel)
Bay and Row work as a "constant" string, Tiers takes the total amount of locations within a Row and uses the value to define the range within a FOR loop.
This creates the labels for the Location in a .docx format in which each page (with the specified dimensions) is a location.
Example:
Bay = "A"
Row = "1"
Tiers = int(10) 
Outcome: A-1-1, A-1-2, A-1-3, ... , A-1-10

If Row and Tiers value == empty:
The Label printer allows you to print the input for Bay, the amount of copies is then chosen once again by the user.
This is useful to create labels for specific items.
