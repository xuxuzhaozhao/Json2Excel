/**
 * 将Json数据导出Excel (暂仅支持非IE浏览器)
 * @param jsonData 样例：{ headers : { 'colName1': '列名1', 'colName2': '列名2', ... }, datalist : [ { 'colName1' : 'content1', 'colName2' : 'content2', ... }, ... ] }
 * @param fileName 导出的文件名称 (不要填写扩展名)，如果没有填写，则默认为‘ecrfplus.xls’
 * @param extension '<tr><td colspan='2'>可自定义最后一行数据<td></tr>' (非必填参数)
 * @example Json2Excel(jsonData, 'ecrfplus')
 * @author xuxuzhaozhao87@gmail.com
 */
function Json2Excel(jsonData, fileName, extension) {
    let excelTemplate = `
		<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
		<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">
		<head>
			<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Sheet1</x:Name><x:WorksheetOptions><x:Selected/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
		</head>
		<body>
			<table border='1' style='font-size:15px;'>{table0}</table>
		</body>
		</html>`;

    const { headers, datalist } = jsonData;
    let headerTitles = Object.keys(headers);
    let headerTemplate = "<tr>";
    headerTitles.forEach(title => {
        headerTemplate += `<th style='text-align:center;'>${headers[title]}</th>`;
    });
    headerTemplate += "</tr>";

    let dataTemplate = "";
    datalist.forEach(data => {
        dataTemplate += "<tr>";
        headerTitles.forEach(title => {
            let currentValue = data[title] || "";
            if (title == 'image') {
                dataTemplate += insertImage(currentValue);
            } else {
                let msoFormat = '@'; // 字符串格式
                let textAlign = "left";
                if (!isNaN(currentValue)) {
                    textAlign = "right";
                    msoFormat = currentValue.toString().indexOf('.') > -1 ? '0\\.00' : '0'; // 数字格式
                }
                dataTemplate += `<td style='mso-number-format:"${msoFormat}"; text-align:${textAlign};' class='tdRight'>${currentValue}</td>`;
            }
        });
        dataTemplate += "</tr>";
    });
    let contentTemplate = headerTemplate + dataTemplate + (extension || "");
    excelTemplate = excelTemplate.replace('{table0}', contentTemplate);

    let blob = new Blob([excelTemplate], { type: "application/vnd.ms-excel" });
    window.URL = window.URL || window.webkitURL;
    let link = window.URL.createObjectURL(blob);
    let a = document.createElement("a");
    a.download = fileName ? fileName : "ecrfplus";
    a.href = link;

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    // 接受http或https作为img的src
    function insertImage(src, width = 40, height = 60) {
        return `<td style="width:${width}px; height:${height}px; text-align:center; vertical-align:middle"><img src="${src}" width=${width} height=${height}></td>`;
    }
}
