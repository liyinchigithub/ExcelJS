// 引入文件、路径、http请求模块
var fs = require('fs');
var path = require('path');
// 发起客户端请求
var request = require('request');
// 引入exceljs模块
var Excel = require('exceljs');
// 实例化一个工作簿对象
var workbook = new Excel.Workbook();

var exc = {
/**
 * @method WriterExcelAndImportFile
 * @description 将客户端请求数据写入到指定的excel工作表中，并将报表导入到指定的地址。
 * @author 李银池
 * @param fileName 
 * @param data 需要写入excel报表的数据
 * @param sheetName 指定写入的excel工作表，req.body.data参数将写到该工作表中
 * @param importUrl 报表进行导入操作的地址
 */
	WriterExcelAndImportFile: (obj, res, next) => {
		/**
   * 
   * 第一部分：
   * @description 写入数据到excel报表指定单元格
   * 使用exceljs
   * 
   */

		//从文件中读取
		workbook.xlsx.readFile(obj.fileName).then(function() {
			// 使用 workbook
			var worksheet = workbook.getWorksheet(obj.sheetName); //根据客户端请求参数obj.sheetName获取指定工作表
			// 根据客户端请求参数obj.data对象，遍历对象每一个key，key就是要写入报表指定位置的单元格位置
			var object = obj.data;
			//遍历设置工作表单元格的值
			for (const key in object) {
				if (object.hasOwnProperty(key)) {
					const element = object[key];
					//通过worksheet.getCell(key).value设置单元格的内容
					worksheet.getCell(key).value = element;
					console.log('key：' + key);
					console.log('element：' + element);
				}
			}

			// 在读取文件回调函数中写入 workbook
			// workbook.xlsx
			// 	.writeFile(obj.fileName)
			// 	.then(function() {
			// 		console.log('数据写入excel成功');
			// 		/**
      //       * 
      //       * 第二部分：
      //       * @description 将数据写入excel完成的文件，执行导入付款结果操作
      //       * 使用request库
      //       * 
      //       */
			// 		//买买车自营结算单导入操作
			// 		var options = {
			// 			//请求地址
			// 			url: obj.importUrl, //可通过req.query属性获取
			// 			//请求头
			// 			headers: {
			// 				'Content-Type': 'multipart/form-data',
			// 				Cookie: obj.cookie_MMC
			// 				//可通过req.headers属性获取
			// 			},
			// 			//请求body
			// 			formData: {
			// 				fileName: '导入付款结果模板.xlsx',
			// 				settlementType: obj.settlementType,
			// 				upfile: {
			// 					value: fs.createReadStream(path.join(__dirname, '/导入付款结果模板.xlsx')),
			// 					options: {
			// 						filename: '导入付款结果模板.xlsx',
			// 						contentType: null
			// 					}
			// 				}
			// 			}, //可通过obj.body属性获取
			// 			json: false
			// 		};
			// 		console.log(path.join(__dirname, '/导入付款结果模板.xlsx'));

			// 		request.post(options, function(error, response, body) {
			// 			console.info('response:' + JSON.stringify(response));
			// 			console.info('statusCode:' + response.statusCode); //获取响应状态码
			// 			console.info('body:' + body);
			// 			if (response.statusCode === 200 && JSON.parse(body).result === true) {
			// 				res.status(200).send('导入付款结果成功');
			// 			} else {
			// 				res.status(200).send('导入付款结果失败');
			// 			}
			// 		});
			// 	})
			// 	.catch(function() {
			// 		console.log('出错了');
			// 	});
		});
	},
	setCell:()=>{
		
	},


};

module.exports = exc;
