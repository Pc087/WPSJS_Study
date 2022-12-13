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
  使用FileSearch 遍历目录下的指定格式的文件
  常用功能如下：
  	如何使用：let FS = Application.FileSearch
  	此时 FS 即为一个搜索对象,以下都用 FS 变量 
  		怎么搜索，通过以下设置
  		1、在哪个目录搜索：FS.LookIn = myPath
  			myPath 可以是一个从 FileDialog选择得到的文件夹名字或者自定义为："D:\\DATA"
  			表示在这个目录下搜索
  		2、是否遍历子目录：FS.SearchSubFolders
  			假如myPath下还有一个文件A，如果想连同子目录内的文件也得到只需要设置为true
  			FS.SearchSubFolders = true
  		3、搜索什么样的文件：FS.FileName = "*.xls"
  			可以使用通配符，如上表示搜索所有.xls结尾的文件
  		4、最后更改时间限制：FS.LastModified = msoLastModifiedAnyTime
  			如上表示所有时间
  			它代表自从上次修改和保存指定的文件的时间量(在文件夹可以看到最后的修改时间)
  			常用枚举值：
  				msoLastModifiedAnyTime		-所有时间 				7
  				msoLastModifiedLastMonth	-最后修改于最近一个月		5
  				msoLastModifiedLastWeek	-最后修改于最近一周		3
  				msoLastModifiedThisMonth	-最后修改于本月			6
  				msoLastModifiedThisWeek	-最后修改于本周			4
  				msoLastModifiedToday		-最后修改于今天			2
  				msoLastModifiedYesterday	-最后修改于昨天			1
  		5、文件类型：FS.FileType
  			可以简单的设定某个文件类型
  			常用枚举值：太多了。。。几乎囊括了常用的所有文件格式
  			下面是几个常用的
  			msoFileTypeDatabases 	-数据库文件 (*.mdb)
  			msoFileTypeOfficeFiles	-文件的任何以下扩展名: *.doc、 .xls，.ppt、 *.pps、 *.obd、 *.mdb、 .mpd，.dot、 .xlt，.pot、 *.obt、 *.htm，或 *.html
  			msoFileTypeWebPages		-HTML 文件 (*.htm 或 *.html)
  			msoFileTypePowerPointPresentations
  									-PowerPoint 演示文稿文件 (.ppt)，PowerPoint 模板文件 (.pot)，或 PowerPoint 幻灯片放映文件 (*.pps)
  			FS.FileType.Add() 方法可往内部传入参数（上述枚举值来添加文件类型)
  		

  ​		6、排序 FS.Execute(msoSortBy, msoSortOrder, Boolean)
  ​			msoSortBy:按照什么来排序
  ​			枚举值：
  ​			msoSortByFileName 	-按文件名
  ​			msoSortByFileType	-按文件类型
  ​			msoSortByLastModified	-最后修改时间
  ​			msoSortBySize		-文件大小
  ​			msoSortOrder:升序或者降序
  ​			msoSortOrderAscending	升序
  ​			msoSortOrderDescending	降序
  ​		7、最后一步：FS.FoundFiles
  ​			将返回符合条件的一个搜索集合对象。
  ​			通过遍历可以取出所有符合条件的文件名
  ​			例如:let files = FS.FoundFiles
  ​			通过遍历1 到 files.Count 取出使用 files(i)

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