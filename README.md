# Json2Excel
> 老项目使用Excel的com组件来导出Excel数据，然而这太占用服务器资源了，所以充分利用Excel可以通过特殊的格式打开Html文件，实现只需要将服务端的JsonData按照特定的格式传递到前端再使用我提供的Json2Excel方法即可导出超大数据量的Excel文件。（经测试，使用COM组件导出20000条数据需要10分钟左右，使用Json2Excel只需10秒左右，且大部分时间是数据库执行复杂SQL语句所费时间）

Example：
```
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
</head>
<body>
    <div>
        导出文件名：<input id="fileName" type="text" />
        导出数据条数：<input id="dataCounts" type="text" />
        <button id="TableExport">TableExportTest</button>
    </div>
    <script src="/IView/Includes/Json2Excel/Json2Excel.js"></script>
    <script>
        var jsonData = {
            "headers": {
                "title1": "标题1",
                "title2": "标题2",
                "title3": "标题3",
                "title4": "标题4",
                "title5": "标题5",
                "title6": "标题6",
                "title7": "标题7",
                "title8":"标题8"
            },
            "datalist": []
        }

        document.getElementById('TableExport').addEventListener("click", () => {
            for (let i = 1; i < Number(document.getElementById('dataCounts').value); i++) {
                jsonData.datalist.push(
                    {
                        "title1": i,
                        "title2": "10.3342",
                        "title3": "内容" + i,
                        "title4": "内容" + i,
                        "title5": "内容" + i,
                        "title6": "内容" + i,
                        "title8": "内容" + i
                    }
                );
            }
            let fileName = document.getElementById('fileName').value;
            let ext = "<tr><td>费用总额：</td><td colspan='7' style='text-align:center'>￥3500.98</td></tr>";
            Json2Excel(jsonData, fileName, ext);
        });    
    </script>
</body>
</html>
```
