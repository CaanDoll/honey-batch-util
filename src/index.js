'use strict';
const fs = require('fs');
const path = require('path');
const xlsx = require('node-xlsx');
// const ejsExcel = require('ejsexcel');
const csvToJson = require('csvtojson');
const root = process.cwd();
const iconv = require('iconv-lite');
const dev = process.env.NODE_ENV === 'development';

let runDir = root;
if (dev) {
  // 测试目录
  runDir = path.join(runDir, 'example');
}
fs.readdir(runDir, (err, files) => {
  if (err) {
    console.error('读取当前运行目录错误');
    console.error(err);
  }
  const generateDir = path.join(runDir, '复评');
  if (!fs.existsSync(generateDir)) {
    fs.mkdirSync(generateDir);
  }
  for (const file of files) {
    if (file.indexOf('.csv') === -1) { // 不是csv文件
      continue;
    }
    const noRepeatArr = []; // 保存去重数组
    const filePath = path.join(runDir, file);
    const stream = fs.createReadStream(filePath, { encoding: 'binary' });
    let data = '';
    stream.on('error', err => {
      console.error('读取行错误');
      console.error(err);
    });
    stream.on('data', chunk => {
      data += chunk;
    });
    stream.on('end', () => {
      // 重新编码，否则乱码
      const buf = Buffer.from(data, 'binary');
      const str = iconv.decode(buf, 'GBK');
      csvToJson({
        noheader: true,
        output: 'csv',
      }).fromString(str)
        .subscribe((csvRow) => {
          if (!noRepeatArr.find(item => item[ 1 ] === csvRow[ 1 ]) && /^http/.test(csvRow[ 4 ])) { // 名称去重 以及 去掉第五列不以http开头的网址
            noRepeatArr.push(csvRow);
          }
        }, err => {
          console.error('读取行错误');
          console.error(err);
        }, () => {
          // 随机排序
          const newArr = noRepeatArr.sort(() => Math.random() - 0.5);
          // 生成文件
          const data = [
            [ 'query', '改进现场', 'url', '打分', '备注' ],
          ];
          for (const item of newArr) {
            data.push([ item[ 1 ], '=HYPERLINK(C2,"现场链接")', item[ 4 ] ]);
          }
          const buffer = xlsx.build([ { name: "sheet1", data } ]);
          const newFile = file.replace('日', '日复评—').replace('.csv', '.xlsx');
          fs.writeFile(path.join(generateDir, newFile), buffer, 'utf8', () => {
            if (err) {
              console.error('写入错误');
              console.error(err);
            }
          })
          /* fs.readFile(path.join(__dirname, 'template.xlsx'), async (err,exlBuf) => {
            if (err) {
              console.error('读取模板错误');
              console.error(err);
            }
            const buffer = await ejsExcel.renderExcel(exlBuf, newArr);
            const newFile = file.replace('日', '日复评—').replace('.csv', '.xlsx');
            fs.writeFile(path.join(generateDir, newFile), buffer, 'utf8', () => {
              if (err) {
                console.error('写入错误');
                console.error(err);
              }
            })
          });*/
        });
    });
  }
});
