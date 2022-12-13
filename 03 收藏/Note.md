# 1、同一工作簿多表汇总

```javascript
function 汇总(){
	const total = Array.from(new Array(Sheets.Count + 1).keys()).slice(1)
		.reduce((prev,cur)=>{
            let arr = Sheets.Item(cur).Range("a1").CurrentRegion.Value2
            if(cur>1) arr.shift();
            return prev.concat(arr)
	},[])
	Sheets.Add().Name = "汇总"
	Range("a1").Resize(total.length,total[0].length).Value2 = total
}
```

# 2、RGB设置颜色

```javascript
function test()
{
	var rgb = (r,g,b)=>(1 << 24) + (b << 16) + (g << 8) + r
	Range("a2").Interior.Color = rgb(255,0,0)
}
```

# 3、导出图片
```javascript
function export_image(){
	ActiveSheet.Shapes(1).SaveAsPicture("D:/Desktop/test.png")
}
```


# 4、Application的方法中参数比较多的简写 Optional Params
- 目前经测试只有application下的方法支持这种传对象参数的写法
```javascript
function test(){
	var RGB = (r,g,b)=>(1 << 24) + (b << 16) + (g << 8) + r
	var Area_rng = Application.InputBox({prompt:"请选择单元格",type:8})
	if(!Area_rng) return
	for(const rng of Area_rng){
		if(rng.Value2<60) rng.Interior.Color = RGB(255,0,0)
	}
}
```

# 5、Array.Sort() 汉字的排序

```javascript
function test(){
	var arr = Range("a1").CurrentRegion.Value2.slice(1).sort((a,b)=>a[3].localeCompare(b[3]))
}
```

# 6、右键菜单

```javascript
function test(){
	Application.CommandBars("Cell").Reset()
	with(Application.CommandBars("Cell").Controls.Add({before:1}))
	{
		FaceId = 2
		Caption = "测试"
		OnAction = "按钮事件"
	}
}
function 按钮事件()
{
	MsgBox("haha")
}
```

