# 1、可变参数

```javascript
function Add(...nums)
{
	// let total=0;
	// nums.forEach(x=>total+=x);
	return nums.reduce((x,y)=>x+=y,0)
}
```

# 2、去重

```javascript
function test(){
    var arr = [1,2,3,2,3,4]
    var res = [...new Set(arr)]
}
```

- 运行结果

```javascript
\\ [1,2,3,4]
```

# 3、文件遍历操作

- Dir遍历

```javascript
function Dir文件遍历(){
	// 路径格式化字符串 String.raw `路径`
	var f_path = String.raw `C:\Users\Administrator\Desktop\test\*.*`
	var i = 0
	var f = Dir(f_path)
	while(true){
		i++
		Console.log(f)
		try{f = Dir()}catch(err){break}
	}
	Console.log(i)
}
```

- Application.FileSearch

```javascript
function test(){
	const fileSearch = Application.FileSearch
	fileSearch.NewSearch()
	fileSearch.LookIn = String.raw `C:\Users\Administrator\Desktop\test`
	fileSearch.FileName = "*.txt"
	const total = fileSearch.Execute() 
	if(total>0){
		for(let i=1;i<=total;i++){
			Console.log(fileSearch.FoundFiles.Item(i))
		}
	}
}
```

# 4、同一工作簿多表汇总

```javascript
function 汇总(){
	const lst = Array.from(new Array(Sheets.Count + 1).keys()).slice(1)
	const total = lst.reduce((prev,cur)=>{
		let arr = Sheets.Item(cur).Range("a1").CurrentRegion.Value2
		if(cur>1) arr.shift();
		return prev.concat(arr)
	},[])
	Sheets.Add()
	ActiveSheet.Name = "汇总"
	Range("a1").Resize(total.length,total[0].length).Value2 = total
}
```

# 5、RGB设置颜色

```javascript
function test()
{
	var rgb = (r,g,b)=>(1 << 24) + (b << 16) + (g << 8) + r
	Range("a2").Interior.Color = rgb(255,0,0)
}
```

<br>

# 6、导出图片
```javascript
function export_image(){
	ActiveSheet.Shapes(1).SaveAsPicture("D:/Desktop/test.png")
}
```


# 7、Application的方法中参数比较多的简写 Optional Params
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
