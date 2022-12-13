# 1、文件遍历操作

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