/**
 * 将Json数据导出Excel (暂仅支持非IE浏览器)
 * @param jsonData 样例：{ headers : [ 'colName1', 'colName2', ... ], dataList : [ { 'colName1' : 'content1', 'colName2' : 'content2', ... }, ... ] }
 * @param fileName 导出的文件名称 (不要填写扩展名)，如果没有填写，则默认为‘ecrfplus.xls’
 * @example Json2Excel(jsonData, 'ecrfplus')
 * @author xucy@ecrfplus.com
 */
function Json2Excel(jsonData, fileName) {
	let excelTemplate = `
		<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
		<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">
		<head>
			<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name/><x:WorksheetOptions><x:Selected/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
		</head>
		<body>
			<table border='1' style='font-size:15px;'>{table0}</table>
		</body>
		</html>`

	const { headers, dataList } = jsonData;
	let headerTemplate = "<tr>"
	headers.forEach(header => {
		headerTemplate += `<th style='text-align:center;'>${header}</th>`;
	});
	headerTemplate += "</tr>";

	let dataTemplate = "";
	dataList.forEach(data => {
		let vals = Object.values(data);
		if (Object.keys(data).length !== headers.length) {
			console.log('错误数据', data);
			alert("数据条数与表头数不一致，检查后重新导出！");
			throw new Error('数据条数与表头数不一致，检查后重新导出！');
		}
		dataTemplate += "<tr>"
		for (let i = 0; i < vals.length; i++) {
			let msoFormat = '@';
			if (!isNaN(vals[i])) {
				msoFormat = vals[i].toString().indexOf('.') > -1 ? '0\.00' : '0';
			}
			dataTemplate += `<td style='mso-number-format:"${msoFormat}"; text-align:right;' class='tdRight'>${vals[i]}</td>`;
		}
		dataTemplate += "</tr>"
	});
	let contentTemplate = headerTemplate + dataTemplate;
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
}