# Json2Excel
> 老项目使用Excel的com组件来导出Excel数据，然而这太占用数据库资源了，所以充分利用Excel可以通过特殊的格式打开Html文件，实现只需要将服务端的JsonData按照特定的格式传递到前端再使用我提供的Json2Excel方法即可导出超大数据量的Excel文件。

Example：
```
<body>
    <div>
        导出文件名：<input id="fileName" type="text" />
        导出数据条数：<input id="dataCounts" type="text" />
        <button id="TableExport">TableExportTest</button>
    </div>
    <script src="/Includes/Json2Excel/Json2Excel.js"></script>
    <script>
        var jsonData = {
            "headers": [
                "标题1",
                "标题2",
                "标题3",
                "标题4",
                "标题5",
                "标题6",
                "标题7",
                "标题8"
            ],
            "dataList": []
        }

        document.getElementById('TableExport').addEventListener("click", () => {
            for (let i = 1; i < Number(document.getElementById('dataCounts').value); i++) {
                jsonData.dataList.push(
                    {
                        "标题1": i,
                        "标题2": "10.2",
                        "标题3": "内容" + i,
                        "标题4": "内容" + i,
                        "标题5": "内容" + i,
                        "标题6": "内容" + i,
                        "标题7": "内容" + i,
                        "标题8": "内容" + i
                    }
                );
            }
            let fileName = document.getElementById('fileName').value;
            Json2Excel(jsonData, fileName);
        });
    </script>
</body>
```
