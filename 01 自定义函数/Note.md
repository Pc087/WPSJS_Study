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

