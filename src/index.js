'use strict';
const fs = require('fs');
const path = require('path');
const xlsx = require('node-xlsx');
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
    if (!/\.csv/.test(file) && !/\.xls(x)?/.test(file)) {
      continue;
    }
    const noRepeatArr = []; // 保存去重数组
    const filePath = path.join(runDir, file);
    // 过滤操作
    const noRepeat = row => {
      if (row && !noRepeatArr.find(item => item[ 1 ] === row[ 1 ]) && /^http/.test(row[ 4 ])) { // 名称去重 以及 去掉第五列不以http开头的网址
        noRepeatArr.push(row);
      }
    };
    // 生成文件函数，传入过滤后的源文件
    const generator = noRepeatArr => {
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
      const newFile = file.replace('日', '日复评—').replace('原表', '').replace(/\.(csv|xls(x)?)$/, '.xlsx');
      fs.writeFile(path.join(generateDir, newFile), buffer, 'utf8', () => {
        if (err) {
          console.error('写入错误');
          console.error(err);
        }
      })
    };
    if (/\.csv/.test(file)) {
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
            console.log(csvRow);
            noRepeat(csvRow);
          }, err => {
            console.error('读取行错误');
            console.error(err);
          }, () => {
            generator(noRepeatArr);
          });
      });
    } else {
      const sheet = xlsx.parse(filePath)[ 0 ].data;
      for (let i = 0; i <= sheet.length; i++) {
        noRepeat(sheet[ i ]);
      }
      generator(noRepeatArr);
    }
  }
});
