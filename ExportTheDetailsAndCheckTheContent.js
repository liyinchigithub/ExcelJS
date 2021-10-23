var express = require("express");
var router = express.Router();
var bodyParser = require('body-parser');
var xlsx = require('node-xlsx');
var fs = require('fs');
var path = require("path");
var request = require('request');
router.use(bodyParser.json());
router.all(function (req, res, next) {
    console.info('------当前时间:', new Date());
    console.info("------请求方法：" + req.method);
    console.info("------请求地址：" + req.path);
});

/**
 * @author 李银池
 */
router.post('/ExportTheDetailsAndCheckTheContent', function (req, res, next) {

   
    let dirPath = path.join(__dirname, "download_path");
    if (!fs.existsSync(dirPath)) {
      fs.mkdirSync(dirPath);
      console.log("文件夹创建成功");
    } else {
      console.log("文件夹已存在");
    }

    let stream = fs.createWriteStream(path.join(dirPath, req.body.fileName + ".xls"));
    let options = {
      url: req.body.downUrl + req.body.fileName + ".xls",
    };
    request(options).pipe(stream).on("close", function (err) {
      console.log("文件[" + req.body.fileName + ".xls]下载完毕");
      var list = xlsx.parse(path.join(dirPath, req.body.fileName + ".xls"));
      console.log(JSON.stringify(list[0].data[1]));
      if (!err === null) {
        res.status(500).send({ error: err })
      } else {
        res.status(200).send({ data: list })
      }
  
    });
  
  });
  


module.exports = router;