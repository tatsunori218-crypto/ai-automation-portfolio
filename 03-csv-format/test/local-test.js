const fs = require('fs'), vm = require('vm');
const dir = process.argv[2] || require("path").join(__dirname, "..");
function parseCsv(t){const rows=[];let row=[],f='',q=false;for(let i=0;i<t.length;i++){const c=t[i];if(q){if(c=='"'){if(t[i+1]=='"'){f+='"';i++}else q=false}else f+=c}else if(c=='"')q=true;else if(c==','){row.push(f);f=''}else if(c=='\n'||c=='\r'){if(c=='\r'&&t[i+1]=='\n')i++;row.push(f);rows.push(row);row=[];f=''}else f+=c}if(f!==''||row.length){row.push(f);rows.push(row)}return rows}
const pad=n=>String(n).padStart(2,'0');
const ctx={Utilities:{parseCsv,formatDate:(d)=>`${d.getFullYear()}/${pad(d.getMonth()+1)}/${pad(d.getDate())}`},Session:{getScriptTimeZone:()=>'Asia/Tokyo'},console};
vm.createContext(ctx); vm.runInContext(fs.readFileSync(dir+'/src/Code.js','utf8'),ctx);
const seen=new Set();
for (const [f,enc] of [['sample_orders_utf8.csv','utf-8'],['sample_orders_sjis.csv','shift_jis'],['sample_orders_add.csv','utf-8']]) {
  const text=new TextDecoder(enc).decode(fs.readFileSync(dir+'/samples/'+f)).replace(/^﻿/,'');
  const r=ctx.formatCsv_(text,seen);
  console.log(`== ${f}: 読込${r.read} 出力${r.rows.length} 重複${r.duplicates} 要確認${r.issueCount}`);
  r.rows.forEach(x=>console.log(x.map(v=>v instanceof Date?ctx.Utilities.formatDate(v):v).join(' | '), x.issues?'['+x.issues+']':''));
}
