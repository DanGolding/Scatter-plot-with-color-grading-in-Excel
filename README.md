# Scatter-plot-with-color-grading-in-Excel
VBA code to colour xy-scatter charts in Excel according to a colour map as well as creating a dynamic colour bar.

#### Examples:

Scatter plot without colouring              |  Scatter plot using a sequential colour map
:------------------------------------------:|:---------------------------------------------:
![](/Images/Sequential_(grey).png?raw=true)  |  ![](/Images/Sequential_(colour).png?raw=true)


Scatter plot without colouring              |  Scatter plot using a diveregent colour map
:------------------------------------------:|:---------------------------------------------:
![](/Images/Divergent_(grey).png?raw=true)  |  ![](/Images/Divergent_(colour).png?raw=true)



All the code is in [ChartColouring.bas](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/ChartColouring.bas). An example spreadsheet is also provided for demonstration. There are 4  steps to colouring your charts:

1. Create a text file containing your colour map
2. Run `MakeMap` to create the colour map in your spreadsheet
2. Run `colourChartSequential` or `colourChartDivergent` to colour your chart
3. Create a colour bar by running `MakeColourBar` and then pasting the colour bar as a linked image.

### Creating the colour map file

The first step is to create a text file containing your colour map. I have provided two example files which you can use, one for [sequential](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/Colour%20Map%20(Sequential).txt) data and one for [divergent](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/Colour%20Map%20(Divergent).txt). This file consists of a single line of space separated hex values in the form #000000. To create your own colour maps, I recommend using [Chroma.js Color Scale Helper](http://gka.github.io/palettes/#colors=#ffffcc,#a1dab4,#41b6c4,#2c7fb8,#253494|steps=256|bez=1|coL=1) with no more than 5 colour 'nodes' specified. I highly recommend using [ColorBrewer2.0](http://colorbrewer2.org/#type=sequential&scheme=YlGnBu&n=5) in order to choose sensible nodes.

![](/Images/chromajs_color_scale_helper.png?raw=true)

### MakeMap
