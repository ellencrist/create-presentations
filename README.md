# Create Presentation with VBA
This is VBA code that creates a PowerPoint presentation from graphs copied from Excel spreadsheets. The code loops through several tabs in Excel, copies the desired charts, and pastes them onto corresponding slides in the presentation.

## Instructions for use

1. Open Excel containing the tabs with the desired graphics.
2. I named the tabs as follows ("A", "B", "C", "D", "E"), change this part of the code according to the name of your tabs.
3. The charts are named according to the default (graphic 1, graphic 2...), but I recommend naming them to facilitate organization, in excel select the chart, click on the format tab, and in the selection panel, and rename the charts any way you like.
5. Open the VBA editor in Excel (press ALT+F11).
6. Insert a new module.
7. Paste the provided code into this module.
8. Define the path and name of the destination file for the PowerPoint presentation by adjusting the path variable.
9. Execute the procedure `CriarPresentation()`.

The presentation will be created with slides corresponding to the graphs in Excel's "A", "B", "C", "D" and "E" tabs. The graphics will be copied and pasted into the presentation slides, according to the position and size specified in the code.

**Note:** Make sure the "Microsoft PowerPoint x.x Object Library" reference is enabled in the VBA editor, where "x.x" represents the installed version of PowerPoint.

<a href="https://uploaddeimagens.com.br/images/004/537/372/full/refer%C3%AAncias_vba_project.png?1688937442"><img src="https://uploaddeimagens.com.br/images/004/537/372/full/refer%C3%AAncias_vba_project.png?1688937442" alt="References VBA" border="0"></a>

## Working
<img src="https://i.ibb.co/C708Vfp/68747470733a2f2f7331312e67696679752e636f6d2f696d616765732f53575448302e676966.gif">

## System Requirements

- Microsoft Excel
- Microsoft PowerPoint

## License
Free to use and modify according to user needs

