# Chart Font

### Setter Font

Use chart title as an example.

```js
Excel.run(function (ctx) { 
	var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
	title.format.font.name = "Calibri";
	title.format.font.size = 12;
	title.format.font.color = "#FF0000";
	title.format.font.italic =  false;
	title.format.font.bold = true;
	title.format.font.underline = false;
	return ctx.sync();
});
```

Set chart title to be Calbri, size 10, bold and in red. 

```js
Excel.run(function (ctx) { 
	var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
	title.format.font.name = "Calibri";
	title.format.font.size = 12;
	title.format.font.color = "#FF0000";
	title.format.font.italic =  false;
	title.format.font.bold = true;
	title.format.font.underline = false;
	return ctx.sync();
});
```
