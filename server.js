const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 3000;
const ADMIN_USER = process.env.ADMIN_USER || 'admin';
const ADMIN_PASS = process.env.ADMIN_PASS || '@Jabar3juara';
const AUTH_SECRET = process.env.AUTH_SECRET || 'Jabar3-auth-secret-change-this-before-production';
const DATA_DIR = path.join(__dirname, 'data');
const SALES_FILE = path.join(DATA_DIR, 'sales.json');
const SALES_META = path.join(DATA_DIR, 'sales_meta.json');
const STOCK_FILE = path.join(DATA_DIR, 'stock.json');
const STOCK_META = path.join(DATA_DIR, 'stock_meta.json');
const DAILY_FILE = path.join(DATA_DIR, 'daily.json');
const DAILY_META = path.join(DATA_DIR, 'daily_meta.json');
const REQUIRED_SALES = ['KABUPATEN','KECAMATAN','PRODUK','AAE','PUD','ALOKASI 1 TAHUN','TOTAL S'];
const REQUIRED_STOCK = ['KABUPATEN','GUDANG LINI III','PPTS','PUD'];
const REQUIRED_DAILY = ['TANGGAL','REALISASI'];
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });
app.use(express.json({limit:'1mb'}));
app.use(express.static(path.join(__dirname,'public')));

function signToken(user){
  const exp = Date.now() + 8*60*60*1000;
  const payload = `${user}.${exp}`;
  const sig = crypto.createHmac('sha256', AUTH_SECRET).update(payload).digest('hex');
  return Buffer.from(`${payload}.${sig}`).toString('base64url');
}
function verifyToken(token){
  try{
    const raw=Buffer.from(token,'base64url').toString();
    const parts=raw.split('.'); if(parts.length!==3) return false;
    const [user,exp,sig]=parts;
    if(!user||!exp||!sig||Number(exp)<Date.now()) return false;
    const expected=crypto.createHmac('sha256',AUTH_SECRET).update(`${user}.${exp}`).digest('hex');
    return crypto.timingSafeEqual(Buffer.from(sig),Buffer.from(expected));
  }catch{return false;}
}
function auth(req,res,next){
  const h=req.headers.authorization||'';
  const token=h.startsWith('Bearer ')?h.slice(7):'';
  if(!verifyToken(token)) return res.status(401).json({error:'Unauthorized'});
  next();
}
function readJson(file,fallback){
  try{return fs.existsSync(file)?JSON.parse(fs.readFileSync(file,'utf8')):fallback;}catch{return fallback;}
}
function cleanHeader(v,i){return (v===null||v===undefined||v==='')?`COL_${i}`:String(v).trim();}
function normalizeNumber(v){if(typeof v==='number') return v;if(v===null||v===undefined||v==='') return 0;return Number(String(v).replace(/[^\d.-]/g,''))||0;}
function parseSheet(buffer, originalName, kind){
  const wb=XLSX.read(buffer,{type:'buffer',cellDates:true});
  const first=wb.SheetNames[0]; const ws=wb.Sheets[first];
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
  if(!rows.length) throw new Error('File Excel kosong.');
  let headers=[]; const seen={};
  rows[0].forEach((x,i)=>{const h=cleanHeader(x,i);seen[h]=(seen[h]||0)+1;headers.push(seen[h]>1?`${h}_${seen[h]}`:h);});
  const required=kind==='sales'?REQUIRED_SALES:kind==='stock'?REQUIRED_STOCK:REQUIRED_DAILY;
  const missing=required.filter(x=>!headers.includes(x));
  if(missing.length) throw new Error('Kolom wajib tidak ditemukan: '+missing.join(', '));
  const records=rows.slice(1).filter(r=>r.some(v=>v!==null&&v!=='' )).map(r=>{const o={};headers.forEach((h,i)=>{const v=r[i];o[h]=v instanceof Date?v.toISOString():v??null;});return o;});
  return {records,meta:{sheet:first,originalName,updatedAt:new Date().toISOString(),rowCount:records.length,headers,kind}};
}
function save(kind,parsed){
  fs.mkdirSync(DATA_DIR,{recursive:true});
  const files = kind==='sales' ? [SALES_FILE, SALES_META] : kind==='stock' ? [STOCK_FILE, STOCK_META] : [DAILY_FILE, DAILY_META];
  fs.writeFileSync(files[0],JSON.stringify(parsed.records));
  fs.writeFileSync(files[1],JSON.stringify(parsed.meta,null,2));
}
function readDataPair(file, metaFile){return {data:readJson(file,[]),meta:readJson(metaFile,{})};}
app.get('/api/data',(req,res)=>{
  res.json({
    sales:readDataPair(SALES_FILE,SALES_META).data.length?readDataPair(SALES_FILE,SALES_META):{data:readJson(path.join(DATA_DIR,'data.json'),[]),meta:readJson(path.join(DATA_DIR,'meta.json'),{})},
    stock:readDataPair(STOCK_FILE,STOCK_META),
    daily:readDataPair(DAILY_FILE,DAILY_META)
  });
});
app.post('/api/login',(req,res)=>{const {username,password}=req.body||{};if(username===ADMIN_USER&&password===ADMIN_PASS)return res.json({token:signToken(username)});res.status(401).json({error:'Username atau password salah'});});
app.post('/api/upload/sales',auth,upload.single('file'),(req,res)=>{try{if(!req.file)throw new Error('File belum dipilih.');const ext=path.extname(req.file.originalname).toLowerCase();if(!['.xlsx','.xls'].includes(ext))throw new Error('Format harus .xlsx atau .xls.');const parsed=parseSheet(req.file.buffer,req.file.originalname,'sales');save('sales',parsed);res.json({ok:true,meta:parsed.meta});}catch(e){res.status(400).json({error:e.message});}});
app.post('/api/upload/stock',auth,upload.single('file'),(req,res)=>{try{if(!req.file)throw new Error('File belum dipilih.');const ext=path.extname(req.file.originalname).toLowerCase();if(!['.xlsx','.xls'].includes(ext))throw new Error('Format harus .xlsx atau .xls.');const parsed=parseSheet(req.file.buffer,req.file.originalname,'stock');save('stock',parsed);res.json({ok:true,meta:parsed.meta});}catch(e){res.status(400).json({error:e.message});}});

app.post('/api/upload/daily',auth,upload.single('file'),(req,res)=>{try{
  if(!req.file)throw new Error('File belum dipilih.');
  const ext=path.extname(req.file.originalname).toLowerCase();
  if(!['.xlsx','.xls'].includes(ext))throw new Error('Format harus .xlsx atau .xls.');
  const parsed=parseSheet(req.file.buffer,req.file.originalname,'daily');
  save('daily',parsed);res.json({ok:true,meta:parsed.meta});
}catch(e){res.status(400).json({error:e.message});}});
app.get('/admin',(req,res)=>res.sendFile(path.join(__dirname,'public','admin.html')));
app.get('*',(req,res)=>res.sendFile(path.join(__dirname,'public','index.html')));
app.listen(PORT,'0.0.0.0',()=>console.log(`Jabar 3 running on :${PORT}`));
