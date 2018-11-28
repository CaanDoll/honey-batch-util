'use strict';
const fs = require('fs');
const path = require('path');
const xlsx = require('node-xlsx');
const csvParse = require('csv-parse');
const root = process.cwd();
const iconv = require('iconv-lite');
const dev = process.env.NODE_ENV === 'development';

let runDir = root;
if (dev) {
  runDir = path.join(runDir, 'example');
}
fs.readdir(runDir, (err, files) => {
  if (err) {
    console.error(err);
  }
  for (const file of files) {
    const filePath = path.join(runDir, file);
    const stream = fs.createReadStream(filePath, { encoding: 'binary' });
    let data = '';
    stream.on('data', chunk => {
      data += chunk;
    });
    stream.on('end', () => {
      const buf = new Buffer(data, 'binary');
      const str = iconv.decode(buf, 'GBK');
      csvParse(str, {
        delimiter: ',',
      }, function(err, output) {
        console.log(err);
        console.log(output);
      })
    });
  }
});
