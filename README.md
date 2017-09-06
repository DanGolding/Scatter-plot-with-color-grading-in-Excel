# Scatter-plot-with-color-grading-in-Excel
VBA code to colour xy-scatter charts in Excel according to a colour map as well as creating a dynamic colour bar.

#### Examples:

Scatter plot without colouring              |  Scatter plot using a sequential colour map
:------------------------------------------:|:---------------------------------------------:
![](/Images/Sequential_(grey).png?raw=true) |  ![](/Images/Sequential_(colour).png?raw=true)


Scatter plot without colouring              |  Scatter plot using a diveregent colour map
:------------------------------------------:|:---------------------------------------------:
![](/Images/Divergent_(grey).png?raw=true)  |  ![](/Images/Divergent_(colour).png?raw=true)



All the code is in [ChartColouring.bas](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/ChartColouring.bas). An example spreadsheet is also provided for demonstration. There are 5 steps to colouring your charts:

1. Import [ChartColouring.bas](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/ChartColouring.bas) into your spreadsheet and save it as a .xlsm file.
2. [Create a text file containing your colour map](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/README.md#creating-the-colour-map-file)
3. [Run `MakeMap`](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/README.md#makemap) to create the colour map in your spreadsheet
4. [Run `colourChartSequential` or `colourChartDivergent`](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/README.md#colourchartsequential-and-colourchartdivergent) to colour your chart
5. [Create a colour bar](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/README.md#makecolourbar) by running `MakeColourBar` and then pasting the colour bar as a linked image.

### [Creating the colour map file](#creating-the-colour-map-file)

The first step is to create a text file containing your colour map. I have provided two example files which you can use, one for [sequential](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/Colour%20Map%20(Sequential).txt) data and one for [divergent](https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel/blob/master/Colour%20Map%20(Divergent).txt). This file consists of a single line of space separated hex values in the form #000000. To create your own colour maps, I recommend using [Chroma.js Color Scale Helper](http://gka.github.io/palettes/#colors=#ffffcc,#a1dab4,#41b6c4,#2c7fb8,#253494|steps=256|bez=1|coL=1) with no more than 5 colour 'nodes' specified. I highly recommend using [ColorBrewer2.0](http://colorbrewer2.org/#type=sequential&scheme=YlGnBu&n=5) in order to choose sensible nodes. If you are interested in the theory of how Chroma.js Color Scale Helper creates colour maps that are perceptual linear (and thus superior to interpolating in RGB or HSV space) read [How To Avoid Equidistant HSV Colors](https://www.vis4.net/blog/posts/avoid-equidistant-hsv-colors/) and [Mastering Multi-hued
Color Scales with Chroma.js](https://www.vis4.net/blog/posts/mastering-multi-hued-color-scales/).

![](/Images/chromajs_color_scale_helper.png?raw=true)

### [MakeMap](#makemap)

The next step is to get that colour map out of the text file and into your spread sheet. Run `MakeMap` to do this. You need to fill in a value for the variable `filename` which is the name of the text file with your colour map in it. It will create a new sheet (named after your colour map text file) containing the hex values of your map in column A, the RGB values in columns B to D and a visualisation of the colour map in column F. The other `sub`s expect the colour map sheet in this format so I recommend you don't make any changes to it besides hiding the sheet. The resulting sheet looks like this:

<p align="center">
  <img src="/Images/Colour_map_example.png" />
</p>

### [`colourChartSequential` and `colourChartDivergent`](#colourchartsequential-and-colourchartdivergent)

The next step is to use the colour map you created with `MakeMap` to colour a scatter chart. Use either `colourChartSequential` or `colourChartDivergent` for this. The colouring is based on data that must be stored in a single column somewhere, and there should be nothing below the data in that column. There are 4 variables, all strings, you need to set in these `sub`s:

 - `sheetData` - the name of the worksheet with your data. Note that by *data*, I mean the data that you are using to define the colouring of the chart. For example if you are plotting average daily temperature vs humidity, then *data* might be the column containing the dates.
 - `dataStartCol` - The column letter where your data are.
 - `dataStartRow` - The row number where your data start. So for example if your data are in a column with a header row, you might set this to `2`
 - `chartName` - The name of the chart you want to colour. I recommend changing the name of the chart from the default names like `Chart 1`.
 
The colouring has only been tested on charts with a single series, however it should be easy to adapt the code to account for a multi-series chart. Just change the line `With sheetData.ChartObjects(chartName).Chart.FullSeriesCollection(1)` and replace the `1` at the end to be the series relevant for your chart.

Lastly, for some reason you sometimes need to ***run the colouring sub twice*** before the colours show.

### [MakeColourBar](#makecolourbar)

The last step is to create the colour bar. This is done by resizing and colouring cells in a worksheet and then merging cells to create the tick marks and the tick labels. The `MakeColourBar` subroutine automatically does all this for you. You just need to specify the name of the new worksheet to create the colour bar in as `name` and the name of the sheet with the colour map (created by `MakeMap`) in as `sheetMap`. If you are interested in what this code is doing, I recommend turning the gridlines back on to get a good visual understanding of which cells are merged and why (i.e. to create the tick marks and the tick mark labels) while simultaneously going through the code.

You then need to fill in the minimum and maximum values for your colour bar in cells `D260` and `D261`. You could use a formula over your *data* range such as `=MIN(...)` for this or hardcode a value. Then copy cells `A1:D258`, navigate to the sheet with your chart on it and paste-special as a *linked image*:

<p align="center">
  <img src="/Images/paste_linked_image.png" />
</p>

Resize the pasted colour bar making sure to keep the aspect ratio constant and position it next to your chart. I like to reduce the size of the *plot area* to make space for the colour bar on the right hand side inside the *chart area*. The tick labels typeface will probably now be too small but because it is a linked image you can just go to the sheet with your colour bar and change the typeface size there iteratively until the tick labels match your charts axis labels in size (~20 is a good starting point). You can do any formatting to the tick mark labels at this point by just using the regular formatting options for cells in Excel. You can also add extra text labels using a textbox, this is how I add the *(Brexit)* label in the example image at the beginning of this document. Note that the resizing / formatting might mess up the aspect ratio of the image so I recommend deleting it and repasting and resizing it.
