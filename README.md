# Scatter-plot-with-color-grading-in-Excel
VBA code to colour xy-scatter charts in Excel according to a colour map as well as creating a dynamic colour bar.

#### Examples:

Scatter plot without colouring              |  Scatter plot using a sequential colour map
:------------------------------------------:|:---------------------------------------------:
![](/Images/Sequential_(grey).png?raw=true) |  ![](/Images/Sequential_(colour).png?raw=true)


Scatter plot without colouring              |  Scatter plot using a diveregent colour map
:------------------------------------------:|:---------------------------------------------:
![](/Images/Divergent_(grey).png?raw=true)  |  ![](/Images/Divergent_(colour).png?raw=true)



All the code is in [ChartColouring.bas](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/ChartColouring.bas). An example spreadsheet is also provided for demonstration. There are 4  steps to colouring your charts:

1. Import [ChartColouring.bas](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/ChartColouring.bas) into your spreadsheet and save it as a .xlsm file.
2. [Create a text file containing your colour map](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/README.md#creating-the-colour-map-file)
3. [Run `MakeMap`](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/README.md#makemap) to create the colour map in your spreadsheet
4. Run `colourChartSequential` or `colourChartDivergent` to colour your chart
5. Create a colour bar by running `MakeColourBar` and then pasting the colour bar as a linked image.

### [Creating the colour map file](#creating-the-colour-map-file)

The first step is to create a text file containing your colour map. I have provided two example files which you can use, one for [sequential](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/Colour%20Map%20(Sequential).txt) data and one for [divergent](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/Colour%20Map%20(Divergent).txt). This file consists of a single line of space separated hex values in the form #000000. To create your own colour maps, I recommend using [Chroma.js Color Scale Helper](http://gka.github.io/palettes/#colors=#ffffcc,#a1dab4,#41b6c4,#2c7fb8,#253494|steps=256|bez=1|coL=1) with no more than 5 colour 'nodes' specified. I highly recommend using [ColorBrewer2.0](http://colorbrewer2.org/#type=sequential&scheme=YlGnBu&n=5) in order to choose sensible nodes. If you are interested in the theory of how Chroma.js Color Scale Helper creates colour maps that are perceptual linear (and thus superior to interpolating in RGB or HSV space) read [How To Avoid Equidistant HSV Colors](https://www.vis4.net/blog/posts/avoid-equidistant-hsv-colors/) and [Mastering Multi-hued
Color Scales with Chroma.js](https://www.vis4.net/blog/posts/mastering-multi-hued-color-scales/).

![](/Images/chromajs_color_scale_helper.png?raw=true)

### [MakeMap](#makemap)

The next step is to get that colour map out of the text file and into your spread sheet. Run `MakeMap` to do this. You need to fill in a value for the variable `name` which is the name of the text file with your colour map in it. It will create a new sheet (named after your colour map text file) containing the hex values of your map in column A, the RGB values in columns B to D and a visualisation of the colour map in column F. The other `sub`s expect the colour map sheet in this format so I recommend you don't make any changes to it besides hiding the sheet. The resulting sheet looks like this:

<p align="center">
  <img src="/Images/Colour_map_example.png" />
</p>
