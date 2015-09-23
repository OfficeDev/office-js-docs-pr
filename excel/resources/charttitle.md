# Chart Title

### Getter and setter

Get the `text` of Chart Title from Chart1.

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var title = chart.title;
title.load('text');
return ctx.sync().then(function() {
		console.log(title.text);
});
```

Set the `text` of Chart Title to "My Chart" and Make it show on top of the chart without overlaying.

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.title.text= "My Chart"; 
chart.title.visible=true;
chart.title.overlay=true;

return ctx.sync().then(function() {
		console.log("Char Title Changed");
});
```
