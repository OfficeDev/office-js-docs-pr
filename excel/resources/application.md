### Getter  
```js
Excel.run(function (ctx) { 
	var application = ctx.workbook.application;
	application.load('calculationMode');
	return ctx.sync().then(function() {
		console.log(application.calculationMode);
	});
});
```

### calculate(calculationType: string)
```js
Excel.run(function (ctx) { 
	ctx.workbook.application.calculate('Full');
	return ctx.sync(); 
}); 
```

