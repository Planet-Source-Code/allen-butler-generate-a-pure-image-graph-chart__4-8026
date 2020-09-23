<div align="center">

## Generate a pure image graph/chart


</div>

### Description

This code snippet demonstrates one easy way to create a pure graphic of a bar chart and line chart combined using the OWC.Chart object(msowc.dll). The msowc.dll is a component distributed with microsoft office. The greatest benefit from this is the fact that you can create almost any type of graph as a pure "GIF". My example code is an easy 'Cut and Paste' cookies cutter code that shows you how to make a pure graphic of a bar and line graph all in one. If you stare at this code long enough, you will figure out how i did everything, it is actually very simple. This is great for any site that needs to display a lot of reporting.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Allen Butler](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/allen-butler.md)
**Level**          |Intermediate
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/allen-butler-generate-a-pure-image-graph-chart__4-8026/archive/master.zip)





### Source Code

```
<%
dim BarVal(8)
dim LineVal(8)
dim Val(8)
BarVal(1) = 4
BarVal(2) = 5
BarVal(3) = 6
BarVal(4) = 4
BarVal(5) = 5
BarVal(6) = 6
BarVal(7) = 8
BarVal(8) = 11
LineVal(1) = 5
LineVal(2) = 4
LineVal(3) = 6
LineVal(4) = 8
LineVal(5) = 9
LineVal(6) = 10
LineVal(7) = 8
LineVal(8) = 9
Val(1) = 1
Val(2) = 2
Val(3) = 3
Val(4) = 4
Val(5) = 4
Val(6) = 6
Val(7) = 7
Val(8) = 8
dim ObjChart
dim ChaCon
dim ObjCharts
dim SerCol
set ObjChart = Server.CreateObject("OWC.Chart")
set ChaCon = ObjChart.Constants
set ObjCharts = ObjChart.Charts.Add
'adding my bar
set SerCol = ObjCharts.SeriesCollection.Add
'adding my line
set SerCol1 = ObjCharts.SeriesCollection.Add
ObjCharts.Type = ChaCon.chChartTypeColumnClustered
		'this is my bar in the graph
		SerCol.Caption = "Estimated Income"
		SerCol.SetData ChaCon.chDimCategories, ChaCon.chDataLiteral, Val
		SerCol.SetData ChaCon.chDimValues, ChaCon.chDataLiteral, BarVal
		'-------------------------------
		'this is my line in the graph
		SerCol1.Caption = "Real Income"
		SerCol1.SetData ChaCon.chDimCategories, ChaCon.chDataLiteral, Val
		SerCol1.SetData ChaCon.chDimValues, ChaCon.chDataLiteral, LineVal
		'change	from the default bar graph to a line
		SerCol1.Type = ChaCon.chChartTypeLine
		'----------------------------
'put a title on this graphic...obviously optional
ObjChart.HasChartSpaceTitle=True
ObjChart.ChartSpaceTitle.Caption = "Earnings Breakdown"
'tell it you want a legend in the graphic
ObjChart.HasChartSpaceLegend = True
ObjChart.ChartSpaceLegend.Position = ChaCon.chLegendPositionRight
'for caching issues on browsers you may want to come up with your own
'naming convention and a way to remove old images....this is a rather
'simple issue with many solutions...if you are still stumped by it
'you can email me and i will show you some of the ways i did this
ImagePath=server.mappath("reports/aspin.gif")
ObjChart.ExportPicture ImagePath,"gif", 400, 200
set ChaCon = nothing
set ObjCharts = nothing
set ObjChart = nothing %>
<img src="/reports/aspin.gif" width="400" height="200">
```

