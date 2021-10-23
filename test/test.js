// 引入exceljs模块
const Excel = require('exceljs');
// 引入
const exc = require('../WriterExcelAndImportFile');

describe('', () => {
	it.skip('', () => {
		// 实例化工作簿对象
		var workbook = new Excel.Workbook();
		// 设置工作簿属性
		workbook.creator = '李银池';

		// 工作簿添加工作表
		var worksheet01 = workbook.addWorksheet('第一个sheet1', {
			headerFooter: { firstHeader: 'Hello Exceljs', firstFooter: 'Hello World' } // 页眉页脚
		});

		// 定义列
		worksheet01.columns = [
			{ header: '编号', key: 'id', width: 15 },
			{ header: '姓名', key: 'name', width: 15, style: { font: { name: 'Arial Black' } } },
			{ header: '地址', key: 'address', width: 15 }
		];
		// 筛选
		worksheet01.autoFilter = 'A1:C1';
		// worksheet01.autoFilter = {
		// 	from: {
		// 		row: 3,
		// 		column: 1
		// 	},
		// 	to: {
		// 		row: 5,
		// 		column: 12
		// 	}
		// };

		// 单元格注释，纯文字笔记
		worksheet01.getCell('A1').note = 'Hello, ExcelJS!';

		// 定义行数据，key根据上一步定义的列
		var data = [
			{
				id: 1,
				name: '李斯',
				address: '泉州'
			},
			{
				id: 2,
				name: '张三',
				address: '厦门'
			},
			{
				id: 3,
				name: '孙立',
				address: '福州'
			},
			{
				id: 4,
				name: '赵武',
				address: '杭州'
			}
		];

		// 添加一列新值
		// worksheet01.getColumn(4).values = [ 5, 6, 7, 8, 9 ];
		// worksheet01.getColumn(5).values = [ '王1', '王2', '王3', '王4', '王5' ];
		// worksheet01.getColumn(6).values = [ '北京', '天津', '上海', '深圳', '广州' ];

		// 添加行数据
		worksheet01.addRows(data);

		// 修改单元格字体
		worksheet01.getCell('A1').font = {
			name: 'Comic Sans MS',
			family: 4,
			size: 16,
			underline: true,
			bold: true
		};
		// 将单元格对齐方式
		worksheet01.getCell('A1').alignment = { vertical: 'top', horizontal: 'left' }; // 设置为左上
		worksheet01.getCell('B1').alignment = { vertical: 'middle', horizontal: 'center' }; // 中间居中
		worksheet01.getCell('C1').alignment = { vertical: 'bottom', horizontal: 'right' }; // 右下

		// 在A1周围设置单个细边框
		worksheet01.getCell('A1').border = {
			top: { style: 'thin' },
			left: { style: 'thin' },
			bottom: { style: 'thin' },
			right: { style: 'thin' }
		};

		// 在A3周围设置双细绿色边框
		worksheet01.getCell('A2').border = {
			top: { style: 'double', color: { argb: 'FF00FF00' } },
			left: { style: 'double', color: { argb: 'FF00FF00' } },
			bottom: { style: 'double', color: { argb: 'FF00FF00' } },
			right: { style: 'double', color: { argb: 'FF00FF00' } }
		};

		// 在A5中设置厚红十字边框
		worksheet01.getCell('A3').border = {
			diagonal: { up: true, down: true, style: 'thick', color: { argb: 'FFFF0000' } }
		};

		//
		// 用红色深色垂直条纹填充A1
		worksheet01.getCell('B1').fill = {
			type: 'pattern',
			pattern: 'darkVertical',
			fgColor: { argb: 'FFFF0000' }
		};

		// 在A2中填充深黄色格子和蓝色背景
		worksheet01.getCell('C1').fill = {
			type: 'pattern',
			pattern: 'darkTrellis',
			fgColor: { argb: 'FFFFFF00' },
			bgColor: { argb: 'FF0000FF' }
		};

		// 从左到右用蓝白蓝渐变填充A3
		worksheet01.getCell('D1').fill = {
			type: 'gradient',
			gradient: 'angle',
			degree: 0,
			stops: [
				{ position: 0, color: { argb: 'FF0000FF' } },
				{ position: 0.5, color: { argb: 'FFFFFFFF' } },
				{ position: 1, color: { argb: 'FF0000FF' } }
			]
		};

		// 从中心开始用红绿色渐变填充A4
		worksheet01.getCell('E1').fill = {
			type: 'gradient',
			gradient: 'path',
			center: { left: 0.5, top: 0.5 },
			stops: [ { position: 0, color: { argb: 'FFFF0000' } }, { position: 1, color: { argb: 'FF00FF00' } } ]
		};

		// 保护单元格
		worksheet01.getCell('A1').protection = {
			locked: false,
			hidden: true
		};
		// 写入到因硬盘中
		workbook.xlsx.writeFile('./file/test2.xlsx').then(
			() => {
				console.log('写入成功');
			},
			(err) => {
				console.log(err);
			}
		);

		// 合并一系列单元格
		worksheet01.mergeCells('A5:B5');

		// 插入行
		// worksheet01.insertRow(1, { id: 11, name: 'John Doe', address: '漳州' });

		const row = worksheet01.getRow(5);
		console.log('row', row);

		// 获取行数
		var rowCount = worksheet01.getRow();
		console.log('rowCount', rowCount);
	});

	it('生成一个工作表，追加数据', async () => {
		// 实例化一个工作簿对象
		const workbook = new Excel.Workbook();
		// 初始化Excel
		workbook.creator = '李银池'; // 作者
		workbook.lastModifiedBy = '李银池'; // 最后一次修改作者
		workbook.created = new Date(); // 创建时间
		workbook.modified = new Date(); // 编辑时间
		let sheet = workbook.addWorksheet('2020-10报表'); // 生成一个工作表sheet

		// 添加列标题并定义列键和宽度
		sheet.columns = [
			// 列名、
			{ header: '创建日期', key: 'create_time', width: 15 },
			{ header: '单号', key: 'id', width: 15 },
			{ header: '电话号码', key: 'phone', width: 15 },
			{ header: '地址', key: 'address', width: 15 }
		];
		// 定义数据（数组）
		const data = [
			{
				create_time: '2020-10-12',
				id: '2020101201',
				phone: '15959950529',
				address: '厦门市'
			}
		];
		// 添加一个行数组
		sheet.addRows(data);
		// 写入工作簿
		await workbook.xlsx.writeFile('file/用户报表.xlsx').then(async () => {}, function(err) {
			console.log(err);
		});
		// 再次添加数据
		const data2 = [
			{
				create_time: '2020-10-11',
				id: '2020101101',
				phone: '15959950530',
				address: '泉州市'
			}
		];
		// Add an array of rows
		sheet.addRows(data2);
		// Add to workbook
		await workbook.xlsx.writeFile('file/用户报表.xlsx').then(async () => {}, function(err) {
			console.log(err);
		});
	});

	it.skip('生成多级表头', async () => {
		// 实例化一个工作簿对象
		const workbook = new Excel.Workbook();
		// 初始化Excel
		workbook.creator = '李银池'; // 作者
		workbook.lastModifiedBy = '李银池'; // 最后一次修改作者
		workbook.created = new Date(); // 创建时间
		workbook.modified = new Date(); // 编辑时间
		let sheet = workbook.addWorksheet('2020-11报表'); // 生成一个工作表sheet

		// 添加表头
		sheet.getRow(1).values = [ '种类', '销量', , , , '店铺' ];
		sheet.getRow(2).values = [ '种类', '2018-05', '2018-06', '2018-07', '2018-08', '店铺' ];

		// 添加数据项定义，与之前不同的是，此时去除header字段
		sheet.columns = [
			{ key: 'category', width: 30 },
			{ key: '2018-05', width: 30 },
			{ key: '2018-06', width: 30 },
			{ key: '2018-07', width: 30 },
			{ key: '2018-08', width: 30 },
			{ key: 'store', width: 30 }
		];
		const data = [
			{
				category: '衣服',
				'2018-05': 300,
				'2018-06': 230,
				'2018-07': 730,
				'2018-08': 630,
				store: '王小二旗舰店'
			},
			{
				category: '零食',
				'2018-05': 672,
				'2018-06': 826,
				'2018-07': 302,
				'2018-08': 389,
				store: '吃吃货'
			}
		];
		sheet.addRows(data);

		// 合并单元格
		sheet.mergeCells(`B1:E1`);
		sheet.mergeCells('A1:A2');
		sheet.mergeCells('F1:F2');

		// 设置每一列样式
		const row = sheet.getRow(1);
		row.eachCell((cell, rowNumber) => {
			sheet.getColumn(rowNumber).alignment = { vertical: 'middle', horizontal: 'center' };
			sheet.getColumn(rowNumber).font = { size: 14, family: 2 };
		});

		// Add to workbook
		await workbook.xlsx.writeFile('file/用户报表.xlsx').then(async () => {}, function(err) {
			console.log(err);
		});
	});

	it.skip('streamed-workbooked', () => {
		var start_time = new Date();
		var workbook = new Excel.stream.xlsx.WorkbookWriter({
			filename: './file/streamed-workbook.xlsx'
		});
		var worksheet = workbook.addWorksheet('Sheet');

		worksheet.columns = [
			{ header: 'id', key: 'id' },
			{ header: 'name', key: 'name' },
			{ header: 'phone', key: 'phone' }
		];

		var data = [
			{
				id: 100,
				name: 'abc',
				phone: '123456789'
			}
		];
		var length = data.length;

		// 当前进度
		var current_num = 0;
		var time_monit = 400;
		var temp_time = Date.now();

		console.log('开始添加数据');
		// 开始添加数据
		for (let i in data) {
			worksheet.addRow(data[i]).commit();
			current_num = i;
			if (Date.now() - temp_time > time_monit) {
				temp_time = Date.now();
				console.log((current_num / length * 100).toFixed(2) + '%');
			}
		}
		console.log('添加数据完毕：', Date.now() - start_time);
		workbook.commit();

		var end_time = new Date();
		var duration = end_time - start_time;

		console.log('用时：' + duration);
		console.log('程序执行完毕');
	});

	it.skip('写入文件', () => {
		// 实例化一个工作簿对象
		const workbook = new Excel.Workbook();
		// 操作文件
		var filename = 'file/test.xlsx';
		var sheetName = 'Sheet1';
		var data = {
			A2: '这是要写入的A2',
			B2: '这是要写入的B2',
			C2: '这是要写入的C2'
		};
		// 使用 workbook
		var worksheet = workbook.getWorksheet(sheetName); //根据客户端请求参数sheetName获取指定工作表
		// 根据客户端请求参数data对象，遍历对象每一个key，key就是要写入报表指定位置的单元格位置 value就是要写入单元格的值
		var object = data;
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
		//从文件中读取
		workbook.xlsx.readFile(filename).then(function() {});
	});

	it.skip('读取文件数据', () => {
		const excelfile = './file/上线计划.xlsx';
		// 实例化一个工作簿对象
		const workbook = new Excel.Workbook();

		workbook.xlsx.readFile(excelfile).then(function() {
			var worksheet = workbook.getWorksheet(1); //获取第一个worksheet

			worksheet.eachRow(function(row, rowNumber) {
				var rowSize = row.cellCount;
				var numValues = row.actualCellCount;
				//console.log("单元格数量/实际数量:"+rowSize+"/"+numValues);
				// cell.type单元格类型：6-公式 ;2-数值；3-字符串
				row.eachCell(function(cell, colNumber) {
					if (cell.type == 6) {
						var value = cell.result;
					} else {
						var value = cell.value;
					}
					console.log('Cell ' + colNumber + ' = ' + cell.type + ' ' + value);
				});
			});
		});
	});

	it.skip('写入excel后进行导入操作', () => {
		var xlsx = '';
		exc.WriterExcelAndImportFile();
	});
});
