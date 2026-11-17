import { Pool, types } from 'pg';
import { getConnectionString } from '@netlify/database';
import * as XLSX from 'xlsx';
import PDFDocument from 'pdfkit';
import JSZip from 'jszip';
import pdfParse from 'pdf-parse';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
const FONT_PATH = fileURLToPath(new URL('./DejaVuSans.ttf', import.meta.url));

// PostgreSQL BIGINT values are safe numeric IDs in this app; keep the UI's number semantics.
types.setTypeParser(20, (v) => Number(v));
const pool = new Pool({ connectionString: getConnectionString(), max: 3, idleTimeoutMillis: 10000, connectionTimeoutMillis: 10000 });
const COOKIE = 'inventory_session';
const json = (data, status=200, headers={}) => Response.json(data, { status, headers });
const err = (message, status=400) => json({ error: message }, status);
const hash = (s) => crypto.createHash('sha256').update(s).digest('hex');
const today = () => new Date().toISOString().slice(0,10);
const num = (v, fallback=0) => { const n=Number(v); return Number.isFinite(n)?n:fallback; };
const text = (v, fallback='') => String(v ?? fallback).trim();
const cookieToken = (req) => (req.headers.get('cookie')||'').split(';').map(x=>x.trim()).find(x=>x.startsWith(COOKIE+'='))?.slice(COOKIE.length+1) || '';
const cookieHeader = (token, maxAge=43200) => `${COOKIE}=${token}; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=${maxAge}`;
async function body(req) { return await req.json().catch(()=>({})); }
async function userFor(req) {
  const token=cookieToken(req); if(!token) return null;
  const r=await pool.query(`SELECT u.id,u.username,u.role,u.active FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token_hash=$1 AND s.expires_at>NOW() AND u.active=TRUE`,[hash(token)]);
  return r.rows[0]||null;
}
async function auth(req) { const u=await userFor(req); return u; }
async function admin(req) { const u=await auth(req); return u?.role==='admin'?u:null; }
async function makeSession(uid) {
  const token=crypto.randomBytes(36).toString('base64url');
  await pool.query(`DELETE FROM sessions WHERE expires_at<=NOW()`);
  await pool.query(`INSERT INTO sessions(token_hash,user_id,expires_at) VALUES($1,$2,NOW()+INTERVAL '12 hours')`,[hash(token),uid]);
  return token;
}
async function passwordHash(password,salt) { return await new Promise((resolve,reject)=>crypto.pbkdf2(password,salt,310000,32,'sha256',(e,b)=>e?reject(e):resolve(b.toString('hex')))); }
async function requireUser(req) { const u=await auth(req); if(!u) throw Object.assign(new Error('Требуется вход в систему'),{status:401}); return u; }
async function audit(client, action, entity, entityId, details='') { await client.query('INSERT INTO audit_log(action,entity,entity_id,details) VALUES($1,$2,$3,$4)',[action,entity,entityId,details]); }
async function fifoDeduct(client, productId, warehouseId, qty, ref, party='', note='FIFO', movementKind='Расход') {
  if(!Number.isFinite(qty) || qty<=0) throw new Error('Количество должно быть больше нуля');
  // Serialize every stock mutation for the same product, including empty-stock adjustments.
  const product=(await client.query('SELECT id FROM products WHERE id=$1 FOR UPDATE',[productId])).rows[0];
  if(!product) throw new Error('Товар не найден');
  const {rows}=await client.query(`SELECT * FROM batches WHERE product_id=$1 AND warehouse_id=$2 AND qty_remaining>0 ORDER BY received_at ASC,id ASC FOR UPDATE`,[productId,warehouseId]);
  const total=rows.reduce((s,b)=>s+Number(b.qty_remaining),0);
  if(qty<=0 || total+1e-8<qty) throw new Error(`Недостаточно остатка. Доступно: ${total}`);
  let left=qty; const used=[];
  for(const b of rows){ if(left<=1e-8) break; const take=Math.min(left,Number(b.qty_remaining)); await client.query('UPDATE batches SET qty_remaining=qty_remaining-$1 WHERE id=$2',[take,b.id]); used.push({batch_id:b.id,qty:take}); left-=take; }
  await client.query(`INSERT INTO movements(kind,product_id,warehouse_id,qty,ref,party,notes) VALUES($1,$2,$3,$4,$5,$6,$7)`,[movementKind,productId,warehouseId,-qty,ref,party,`${note}; партии: ${JSON.stringify(used)}`]);
  return used;
}
function taxParts(qty, price, rate, mode='inclusive'){ const base=Number(qty||0)*Number(price||0), r=Math.max(0,Number(rate||0)); if(mode==='none') return {net:base,tax:0,gross:base}; if(mode==='exclusive'){const tax=base*r/100;return {net:base,tax,gross:base+tax};} const tax=r?base*r/(100+r):0;return {net:base-tax,tax,gross:base};}
function safeFilename(s){return String(s||'file').replace(/[\\/:*?"<>|\r\n]/g,'_').slice(0,120)}
const BACKUP_TABLES=['manufacturers','suppliers','customers','products','certificates','product_certificates','certificate_catalog_items','certificate_files','warehouses','batches','movements','documents','document_items','audit_log','document_allocations'];
const BACKUP_ID_TABLES=['manufacturers','suppliers','customers','products','certificates','certificate_catalog_items','certificate_files','warehouses','batches','movements','documents','audit_log'];
async function ensureBackupUploadTables(){
  await pool.query(`CREATE TABLE IF NOT EXISTS backup_restore_uploads(upload_id TEXT PRIMARY KEY, created_by BIGINT NOT NULL REFERENCES users(id) ON DELETE CASCADE, file_name TEXT NOT NULL, total_chunks INTEGER NOT NULL, file_size BIGINT NOT NULL, created_at TIMESTAMP NOT NULL DEFAULT NOW())`);
  await pool.query(`CREATE TABLE IF NOT EXISTS backup_restore_chunks(upload_id TEXT NOT NULL REFERENCES backup_restore_uploads(upload_id) ON DELETE CASCADE, chunk_index INTEGER NOT NULL, chunk_data BYTEA NOT NULL, PRIMARY KEY(upload_id,chunk_index))`);
}
async function restoreBackupZip(zipBuffer, client){
  const zip=await JSZip.loadAsync(zipBuffer,{checkCRC32:true});
  const data={};let total=0;
  for(const table of BACKUP_TABLES){
    const entry=zip.file(`database/${table}.json`);if(!entry)throw new Error(`В архиве отсутствует файл database/${table}.json`);
    const raw=await entry.async('string');total+=Buffer.byteLength(raw);if(total>250*1024*1024)throw new Error('Распакованный архив слишком большой (лимит 250 МБ)');
    const rows=JSON.parse(raw);if(!Array.isArray(rows))throw new Error(`Некорректный формат таблицы ${table}`);
    if(rows.length>1000000)throw new Error(`Слишком много записей в таблице ${table}`);
    data[table]=rows;
  }
  // Preserve user accounts and sessions so the administrator can sign in after restore.
  await client.query(`TRUNCATE TABLE ${BACKUP_TABLES.join(',')} RESTART IDENTITY CASCADE`);
  for(const table of ['manufacturers','suppliers','customers','warehouses','products','certificates','product_certificates','certificate_catalog_items','certificate_files','batches','movements','documents','document_items','document_allocations','audit_log']){
    const rows=data[table];
    for(const row of rows){
      if(!row||typeof row!=='object'||Array.isArray(row))throw new Error(`Некорректная запись в ${table}`);
      const cols=Object.keys(row);if(!cols.length)continue;
      if(cols.some(c=>!/^[a-z_][a-z0-9_]*$/i.test(c)))throw new Error(`Недопустимое поле в таблице ${table}`);
      const vals=cols.map(c=>{
        const v=row[c];
        if(table==='certificate_files'&&c==='file_data'&&typeof v==='string')return Buffer.from(v,'base64');
        if(v&&typeof v==='object'&&v.type==='Buffer'&&Array.isArray(v.data))return Buffer.from(v.data);
        return v;
      });
      const placeholders=vals.map((_,i)=>`$${i+1}`).join(',');
      await client.query(`INSERT INTO ${table} (${cols.map(c=>`"${c}"`).join(',')}) VALUES (${placeholders})`,vals);
    }
  }
  for(const table of BACKUP_ID_TABLES){
    await client.query(`SELECT setval(pg_get_serial_sequence('public.${table}','id'),COALESCE(MAX(id),1),MAX(id) IS NOT NULL) FROM ${table}`);
  }
  return Object.fromEntries(BACKUP_TABLES.map(t=>[t,data[t].length]));
}

async function exportXlsx(doc,items){
  const rows=[['ТОВАРНЫЙ ДОКУМЕНТ'],['Номер',doc.doc_no,'Дата',doc.doc_date],['Покупатель',doc.customer||''],['УНП',doc.unp||''],['Адрес',doc.customer_address||''],[],['№','Артикул','Наименование','Код ТН ВЭД','Производитель','Сертификат/декларация','Кол-во','Цена BYN','НДС %','Сумма без НДС','НДС','Итого']];
  const mode=doc.vat_mode||'inclusive'; for(let i=0;i<items.length;i++){const x=items[i],t=taxParts(x.qty,x.price,x.vat,mode);rows.push([i+1,x.article,x.name,x.tnved,x.manufacturer,x.certificate,num(x.qty),num(x.price),mode==='none'?'Без НДС':`${num(x.vat)}%`,t.net,t.tax,t.gross]);}
  const sums=items.reduce((a,x)=>{const t=taxParts(x.qty,x.price,x.vat,mode);a.net+=t.net;a.tax+=t.tax;a.gross+=t.gross;return a;},{net:0,tax:0,gross:0});rows.push(['','','','','','','','','ИТОГО',sums.net,sums.tax,sums.gross]); rows.splice(3,0,['Режим НДС',mode==='inclusive'?'Цена включает НДС':mode==='exclusive'?'НДС начисляется сверху':'Без НДС']);
  const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=[6,20,38,18,30,24,12,14,10,16,14,15].map(w=>({wch:w}));ws['!autofilter']={ref:`A8:L${8+items.length}`};
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Документ');return XLSX.write(wb,{type:'buffer',bookType:'xlsx'});
}
async function exportPdf(doc,items){
  return await new Promise((resolve,reject)=>{const p=new PDFDocument({size:'A4',layout:'landscape',margin:24});const chunks=[];p.on('data',x=>chunks.push(x));p.on('end',()=>resolve(Buffer.concat(chunks)));p.on('error',reject);
    const mode=doc.vat_mode||'inclusive',modeLabel=mode==='inclusive'?'Цена включает НДС':mode==='exclusive'?'НДС начисляется сверху':'Без НДС';
    p.font(FONT_PATH).fontSize(16).text('ТОВАРНЫЙ ДОКУМЕНТ',{align:'center'});p.moveDown(.4).fontSize(8).text(`Номер: ${doc.doc_no||''}    Дата: ${doc.doc_date||''}`);p.text(`Покупатель: ${doc.customer||''}    УНП: ${doc.unp||''}`);p.text(`Адрес: ${doc.customer_address||''}`);p.text(`Налоговый режим: ${modeLabel}`);p.moveDown();
    const headers=['№','Артикул','Наименование','ТН ВЭД','Производитель','Кол-во','Цена','Ставка','НДС','Итого'];const widths=[22,66,180,66,90,42,52,40,50,58];let y=p.y;
    const drawRow=(vals,bold=false)=>{let x=24;const h=bold?23:31;vals.forEach((v,i)=>{p.font(FONT_PATH).fontSize(6.6).text(String(v??''),x+2,y+3,{width:widths[i]-4,height:h-4,ellipsis:true});x+=widths[i];});p.moveTo(24,y+h).lineTo(24+widths.reduce((a,b)=>a+b,0),y+h).strokeColor('#bbbbbb').stroke();y+=h;};
    drawRow(headers,true);let sums={net:0,tax:0,gross:0};items.forEach((x,i)=>{const t=taxParts(x.qty,x.price,x.vat,mode);sums.net+=t.net;sums.tax+=t.tax;sums.gross+=t.gross;drawRow([i+1,x.article,x.name,x.tnved,x.manufacturer,x.qty,Number(x.price).toFixed(2),mode==='none'?'—':`${x.vat}%`,t.tax.toFixed(2),t.gross.toFixed(2)]);if(y>520){p.addPage();y=28;drawRow(headers,true);}});y+=8;p.font(FONT_PATH).fontSize(9).text(`Без НДС: ${sums.net.toFixed(2)} BYN    НДС: ${sums.tax.toFixed(2)} BYN    ИТОГО: ${sums.gross.toFixed(2)} BYN`,24,y);p.end();});
}
async function docWithItems(id){const d=(await pool.query(`SELECT d.*,c.name customer,c.unp,c.address customer_address,w.name warehouse FROM documents d LEFT JOIN customers c ON c.id=d.customer_id LEFT JOIN warehouses w ON w.id=d.warehouse_id WHERE d.id=$1`,[id])).rows[0];if(!d)return null;const items=(await pool.query('SELECT * FROM document_items WHERE document_id=$1 ORDER BY id',[id])).rows;return {d,items};}
async function handler(req){
  const url=new URL(req.url);let path=url.pathname.replace(/^\/\.netlify\/functions\/api/,'/api');if(!path.startsWith('/'))path='/'+path;const method=req.method.toUpperCase();
  // Public authentication endpoints
  if(path==='/api/auth/status'&&method==='GET'){
    const count=Number((await pool.query('SELECT COUNT(*)::int n FROM users')).rows[0].n);const u=await userFor(req);return json({setup_required:count===0,authenticated:!!u,user:u});
  }
  if(path==='/api/auth/setup'&&method==='POST'){
    const d=await body(req);const username=text(d.username,'admin'),password=String(d.password||'');if(username.length<3||password.length<12)return err('Логин — минимум 3 символа, пароль — минимум 12');
    const c=await pool.connect();try{await c.query('BEGIN');await c.query('SELECT pg_advisory_xact_lock(7291042026)');const n=Number((await c.query('SELECT COUNT(*)::int n FROM users')).rows[0].n);if(n)throw new Error('Первичная настройка уже выполнена');const salt=crypto.randomBytes(16).toString('hex');const ph=await passwordHash(password,salt);const u=(await c.query(`INSERT INTO users(username,password_hash,salt,role) VALUES($1,$2,$3,'admin') RETURNING id`,[username,ph,salt])).rows[0];await c.query('COMMIT');const token=await makeSession(u.id);return json({ok:true},200,{'Set-Cookie':cookieHeader(token)});}catch(e){await c.query('ROLLBACK');return err(e.message,400)}finally{c.release()}
  }
  if(path==='/api/auth/login'&&method==='POST'){
    const d=await body(req),username=text(d.username),password=String(d.password||'');const u=(await pool.query('SELECT * FROM users WHERE username=$1 AND active=TRUE',[username])).rows[0];if(!u)return err('Неверный логин или пароль',401);const ph=await passwordHash(password,u.salt);if(!crypto.timingSafeEqual(Buffer.from(ph,'hex'),Buffer.from(u.password_hash,'hex')))return err('Неверный логин или пароль',401);const token=await makeSession(u.id);return json({ok:true},200,{'Set-Cookie':cookieHeader(token)});
  }
  if(path==='/api/auth/logout'&&method==='POST'){const t=cookieToken(req);if(t)await pool.query('DELETE FROM sessions WHERE token_hash=$1',[hash(t)]);return json({ok:true},200,{'Set-Cookie':cookieHeader('',0)});}
  let user;try{user=await requireUser(req)}catch(e){return err(e.message,e.status||401)}
  if(path==='/api/users'&&method==='GET'){if(user.role!=='admin')return err('Недостаточно прав',403);return json((await pool.query('SELECT id,username,role,active,created_at FROM users ORDER BY id')).rows)}
  if(path==='/api/users'&&method==='POST'){if(user.role!=='admin')return err('Недостаточно прав',403);const d=await body(req);if(text(d.username).length<3||String(d.password||'').length<12||!['admin','operator'].includes(d.role||'operator'))return err('Проверьте логин, пароль (минимум 12 символов) и роль');const c=await pool.connect();try{await c.query('BEGIN');await c.query('SELECT pg_advisory_xact_lock(7291042027)');const n=Number((await c.query('SELECT COUNT(*)::int n FROM users')).rows[0].n);if(n>=3){await c.query('ROLLBACK');return err('Можно создать не более трёх пользователей')}const salt=crypto.randomBytes(16).toString('hex'),ph=await passwordHash(String(d.password),salt);const r=await c.query('INSERT INTO users(username,password_hash,salt,role) VALUES($1,$2,$3,$4) RETURNING id',[text(d.username),ph,salt,d.role||'operator']);await c.query('COMMIT');return json({id:r.rows[0].id});}catch(e){await c.query('ROLLBACK');return err(e.message)}finally{c.release()}}
  if(path==='/api/backup'&&method==='GET'){
    if(user.role!=='admin')return err('Резервную копию может скачивать только администратор',403);
    const zip=new JSZip();zip.file('manifest.json',JSON.stringify({app:'mebel-inventory',format_version:1,created_at:new Date().toISOString(),tables:BACKUP_TABLES},null,2));
    for(const t of BACKUP_TABLES){const rows=(await pool.query(`SELECT * FROM ${t}`)).rows;zip.file(`database/${t}.json`,JSON.stringify(rows,(k,v)=>v && v.type==='Buffer' && Array.isArray(v.data) ? Buffer.from(v.data).toString('base64') : v,2));}
    const data=await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'});return new Response(data,{headers:{'Content-Type':'application/zip','Content-Disposition':'attachment; filename="inventory_backup.zip"','Cache-Control':'no-store'}});
  }
  if(path==='/api/backup/restore/start'&&method==='POST'){
    if(user.role!=='admin')return err('Восстанавливать резервную копию может только администратор',403);
    await ensureBackupUploadTables();const d=await body(req),chunks=Number(d.total_chunks),size=Number(d.file_size),name=text(d.file_name,'backup.zip');
    if(!name.toLowerCase().endsWith('.zip'))return err('Выберите ZIP-файл резервной копии');
    if(!Number.isInteger(chunks)||chunks<1||chunks>256||!Number.isFinite(size)||size<1||size>100*1024*1024)return err('Размер архива не должен превышать 100 МБ');
    await pool.query(`DELETE FROM backup_restore_uploads WHERE created_at<NOW()-INTERVAL '2 hours'`);
    const id=crypto.randomUUID();await pool.query('INSERT INTO backup_restore_uploads(upload_id,created_by,file_name,total_chunks,file_size) VALUES($1,$2,$3,$4,$5)',[id,user.id,name,chunks,size]);return json({upload_id:id,chunk_size:1024*1024});
  }
  if(path==='/api/backup/restore/chunk'&&method==='POST'){
    if(user.role!=='admin')return err('Недостаточно прав',403);await ensureBackupUploadTables();const id=url.searchParams.get('id')||'',index=Number(url.searchParams.get('index'));
    const up=(await pool.query('SELECT * FROM backup_restore_uploads WHERE upload_id=$1 AND created_by=$2',[id,user.id])).rows[0];if(!up)return err('Загрузка не найдена или истекла',404);
    if(!Number.isInteger(index)||index<0||index>=up.total_chunks)return err('Неверный номер части');const bytes=Buffer.from(await req.arrayBuffer());if(!bytes.length||bytes.length>1024*1024)return err('Размер части должен быть от 1 байта до 1 МБ');
    await pool.query('INSERT INTO backup_restore_chunks(upload_id,chunk_index,chunk_data) VALUES($1,$2,$3) ON CONFLICT(upload_id,chunk_index) DO UPDATE SET chunk_data=EXCLUDED.chunk_data',[id,index,bytes]);return json({ok:true,index});
  }
  if(path==='/api/backup/restore/finish'&&method==='POST'){
    if(user.role!=='admin')return err('Недостаточно прав',403);await ensureBackupUploadTables();const d=await body(req),id=text(d.upload_id);
    const up=(await pool.query('SELECT * FROM backup_restore_uploads WHERE upload_id=$1 AND created_by=$2',[id,user.id])).rows[0];if(!up)return err('Загрузка не найдена или истекла',404);
    const parts=(await pool.query('SELECT chunk_index,chunk_data FROM backup_restore_chunks WHERE upload_id=$1 ORDER BY chunk_index',[id])).rows;
    if(parts.length!==up.total_chunks||parts.some((p,i)=>Number(p.chunk_index)!==i))return err('Не все части архива загружены. Повторите загрузку.');
    const archive=Buffer.concat(parts.map(p=>p.chunk_data));if(archive.length!==Number(up.file_size))return err('Размер загруженного файла не совпадает с исходным');
    const client=await pool.connect();try{await client.query('BEGIN');const counts=await restoreBackupZip(archive,client);await client.query('COMMIT');await pool.query('DELETE FROM backup_restore_uploads WHERE upload_id=$1',[id]);return json({ok:true,restored:counts});}catch(e){await client.query('ROLLBACK');return err(`Восстановление отменено: ${e.message}`,400)}finally{client.release()}
  }
  if(path==='/api/dashboard'&&method==='GET'){
    const r=await pool.query(`SELECT (SELECT COUNT(*) FROM products WHERE active=TRUE)::int products,(SELECT COUNT(*) FROM customers)::int customers,(SELECT COUNT(*) FROM certificates)::int certificates,(SELECT COALESCE(SUM(qty_remaining),0) FROM batches)::float stock_units,(SELECT COUNT(*) FROM certificates WHERE expiry_date<>'' AND expiry_date<=TO_CHAR(CURRENT_DATE+INTERVAL '60 days','YYYY-MM-DD'))::int expiring,(SELECT COUNT(*) FROM products p WHERE p.active=TRUE AND COALESCE((SELECT SUM(b.qty_remaining) FROM batches b WHERE b.product_id=p.id),0)<=0)::int low_stock`);
    const recent=(await pool.query(`SELECT m.*,p.article,p.name product,w.name warehouse FROM movements m JOIN products p ON p.id=m.product_id JOIN warehouses w ON w.id=m.warehouse_id ORDER BY m.id DESC LIMIT 8`)).rows;return json({...r.rows[0],recent});
  }
  const fileMatch=path.match(/^\/api\/certificate-files\/(\d+)$/)||path.match(/^\/certificate-files\/(\d+)$/);
  if(fileMatch&&method==='GET'){const f=(await pool.query('SELECT * FROM certificate_files WHERE id=$1',[Number(fileMatch[1])])).rows[0];if(!f)return err('Файл не найден',404);if(!f.file_data)return err('Файл не сохранён в базе',404);return new Response(f.file_data,{headers:{'Content-Type':'application/pdf','Content-Disposition':`inline; filename="${safeFilename(f.original_name)}"`}})}
  const exp=path.match(/^\/api\/documents\/(\d+)\/export$/);
  if(exp&&method==='GET'){const x=await docWithItems(Number(exp[1]));if(!x)return err('Документ не найден',404);const fmt=url.searchParams.get('format')||'xlsx';if(fmt==='pdf'){const b=await exportPdf(x.d,x.items);return new Response(b,{headers:{'Content-Type':'application/pdf','Content-Disposition':`attachment; filename="document_${exp[1]}.pdf"`}})}const b=await exportXlsx(x.d,x.items);return new Response(b,{headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','Content-Disposition':`attachment; filename="document_${exp[1]}.xlsx"`}})}
  if(method==='GET'){
    if(path==='/api/products'){const q=url.searchParams.get('q')||'';const r=await pool.query(`SELECT p.*,m.name manufacturer,COALESCE((SELECT SUM(qty_remaining) FROM batches b WHERE b.product_id=p.id),0)::float stock,(SELECT string_agg(ce.number,', ') FROM product_certificates pc JOIN certificates ce ON ce.id=pc.certificate_id WHERE pc.product_id=p.id) certificates FROM products p LEFT JOIN manufacturers m ON m.id=p.manufacturer_id WHERE p.active=TRUE AND ($1='' OR p.article ILIKE '%'||$1||'%' OR p.name ILIKE '%'||$1||'%' OR p.tnved ILIKE '%'||$1||'%' OR m.name ILIKE '%'||$1||'%' OR EXISTS(SELECT 1 FROM product_certificates pc JOIN certificates ce ON ce.id=pc.certificate_id WHERE pc.product_id=p.id AND ce.number ILIKE '%'||$1||'%')) ORDER BY p.article`,[q]);return json(r.rows)}
    if(path==='/api/products/archive'){const q=url.searchParams.get('q')||'';const r=await pool.query(`SELECT p.*,m.name manufacturer,COALESCE((SELECT SUM(qty_remaining) FROM batches b WHERE b.product_id=p.id),0)::float stock FROM products p LEFT JOIN manufacturers m ON m.id=p.manufacturer_id WHERE p.active=FALSE AND ($1='' OR p.article ILIKE '%'||$1||'%' OR p.name ILIKE '%'||$1||'%' OR p.tnved ILIKE '%'||$1||'%' OR m.name ILIKE '%'||$1||'%') ORDER BY p.article`,[q]);return json(r.rows)}
    if(path==='/api/customers'||path==='/api/suppliers'||path==='/api/manufacturers'||path==='/api/warehouses') {const t=path.split('/').pop();return json((await pool.query(`SELECT * FROM ${t} ORDER BY name`)).rows)}
    if(path==='/api/certificates')return json((await pool.query(`SELECT ce.*,m.name manufacturer,CASE WHEN ce.expiry_date<>'' AND ce.expiry_date<TO_CHAR(CURRENT_DATE,'YYYY-MM-DD') THEN 'Истёк' ELSE ce.status END effective_status FROM certificates ce LEFT JOIN manufacturers m ON m.id=ce.manufacturer_id ORDER BY ce.expiry_date`)).rows);
    if(path==='/api/certificate-catalog'){const q=url.searchParams.get('q')||'';return json((await pool.query(`SELECT ci.*,p.article linked_article,p.name linked_product,CASE WHEN ci.expiry_date<>'' AND ci.expiry_date<TO_CHAR(CURRENT_DATE,'YYYY-MM-DD') THEN 'Истёк' ELSE ci.review_status END effective_status FROM certificate_catalog_items ci LEFT JOIN products p ON p.id=ci.product_id WHERE ($1='' OR ci.certificate_number ILIKE '%'||$1||'%' OR ci.product_code ILIKE '%'||$1||'%' OR ci.description ILIKE '%'||$1||'%' OR ci.tnved ILIKE '%'||$1||'%') ORDER BY ci.certificate_number,ci.product_code`,[q])).rows)}
    if(path==='/api/certificate-files')return json((await pool.query(`SELECT f.id,f.certificate_id,f.original_name,f.uploaded_at,ce.number certificate_number FROM certificate_files f LEFT JOIN certificates ce ON ce.id=f.certificate_id ORDER BY f.id DESC`)).rows);
    if(path==='/api/stock')return json((await pool.query(`SELECT b.*,p.article,p.name product,w.name warehouse,s.name supplier,ce.number certificate FROM batches b JOIN products p ON p.id=b.product_id JOIN warehouses w ON w.id=b.warehouse_id LEFT JOIN suppliers s ON s.id=b.supplier_id LEFT JOIN certificates ce ON ce.id=b.certificate_id ORDER BY b.id DESC`)).rows);
    if(path==='/api/movements')return json((await pool.query(`SELECT m.*,p.article,p.name product,w.name warehouse FROM movements m JOIN products p ON p.id=m.product_id JOIN warehouses w ON w.id=m.warehouse_id ORDER BY m.id DESC LIMIT 300`)).rows);
    if(path==='/api/documents')return json((await pool.query(`SELECT d.*,cu.name customer,w.name warehouse FROM documents d LEFT JOIN customers cu ON cu.id=d.customer_id LEFT JOIN warehouses w ON w.id=d.warehouse_id ORDER BY d.id DESC`)).rows);
    if(path==='/api/document-products')return json((await pool.query(`SELECT p.id,p.article,p.name,p.tnved,p.unit,p.vat,p.sale_price,p.sale_price_without_vat,m.name manufacturer,string_agg(ce.number,', ') certificates FROM products p LEFT JOIN manufacturers m ON m.id=p.manufacturer_id LEFT JOIN product_certificates pc ON pc.product_id=p.id LEFT JOIN certificates ce ON ce.id=pc.certificate_id WHERE p.active=TRUE GROUP BY p.id,m.name ORDER BY p.article`)).rows);
  }
  if(method==='POST' && path==='/api/import/products-json'){
    const d=await body(req), rows=Array.isArray(d.rows)?d.rows:[];
    if(!rows.length) return err('В файле не найдены товарные строки');
    if(rows.length>10000) return err('За один импорт можно загрузить не более 10 000 строк');
    let added=0,updated=0,skipped=0;
    const c=await pool.connect();
    try{
      await c.query('BEGIN');
      for(const row of rows){
        const article=text(row.article), name=text(row.name);
        if(!article||!name){skipped++;continue;}
        const mn=text(row.manufacturer);let mid=null;
        if(mn){const mr=await c.query('INSERT INTO manufacturers(name) VALUES($1) ON CONFLICT(name) DO UPDATE SET name=EXCLUDED.name RETURNING id',[mn]);mid=mr.rows[0].id;}
        const exists=(await c.query('SELECT id FROM products WHERE article=$1',[article])).rows[0];
        await c.query(`INSERT INTO products(article,name,category,tnved,manufacturer_id,unit,vat,purchase_price,sale_price,sale_price_without_vat,active)
          VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)
          ON CONFLICT(article) DO UPDATE SET name=EXCLUDED.name,category=EXCLUDED.category,tnved=CASE WHEN EXCLUDED.tnved<>'' THEN EXCLUDED.tnved ELSE products.tnved END,
          manufacturer_id=COALESCE(EXCLUDED.manufacturer_id,products.manufacturer_id),unit=EXCLUDED.unit,vat=EXCLUDED.vat,
          purchase_price=EXCLUDED.purchase_price,sale_price=EXCLUDED.sale_price,sale_price_without_vat=EXCLUDED.sale_price_without_vat,active=EXCLUDED.active`,
          [article,name,text(row.category),text(row.tnved),mid,text(row.unit,'шт'),num(row.vat,20),num(row.purchase_price),num(row.sale_price),num(row.sale_price_without_vat),row.active!==false]);
        if(exists)updated++;else added++;
      }
      await c.query('COMMIT');return json({added,updated,skipped,total:rows.length});
    }catch(e){await c.query('ROLLBACK');throw e}finally{c.release()}
  }
  if(method==='POST' && path==='/api/products/bulk'){
    const d=await body(req), ids=[...new Set((Array.isArray(d.ids)?d.ids:[]).map(Number).filter(x=>Number.isInteger(x)&&x>0))], action=text(d.action);
    if(!ids.length)return err('Выберите хотя бы один товар'); if(ids.length>2000)return err('За одну операцию можно обработать не более 2000 товаров'); if(!['archive','delete'].includes(action))return err('Неизвестное действие');
    const c=await pool.connect();try{await c.query('BEGIN');let archived=0,deleted=0,missing=0;
      for(const id of ids){const exists=(await c.query('SELECT id FROM products WHERE id=$1 AND active=TRUE FOR UPDATE',[id])).rows[0];if(!exists){missing++;continue;}
        const used=(await c.query(`SELECT EXISTS(SELECT 1 FROM batches WHERE product_id=$1) OR EXISTS(SELECT 1 FROM movements WHERE product_id=$1) OR EXISTS(SELECT 1 FROM document_items WHERE product_id=$1) OR EXISTS(SELECT 1 FROM certificate_catalog_items WHERE product_id=$1) AS used`,[id])).rows[0].used;
        if(action==='archive'||used){await c.query('UPDATE products SET active=FALSE WHERE id=$1',[id]);archived++;}
        else {await c.query('DELETE FROM products WHERE id=$1',[id]);deleted++;}
      }
      await c.query('COMMIT');return json({ok:true,archived,deleted,missing,total:ids.length});
    }catch(e){await c.query('ROLLBACK');return err(e.message,409)}finally{c.release()}
  }
  if(method==='POST'||method==='PUT'||method==='PATCH'){
    // Multipart uploads
    if(path==='/api/certificate-files'&&method==='POST'){
      const fd=await req.formData();const file=fd.get('file');if(!file||typeof file.arrayBuffer!=='function')return err('PDF-файл не найден');if(file.size>5*1024*1024)return err('Максимальный размер PDF — 5 МБ');if(file.type!=='application/pdf'&&!file.name.toLowerCase().endsWith('.pdf'))return err('Можно загрузить только PDF');const buf=Buffer.from(await file.arrayBuffer());if(buf.length<5||buf.subarray(0,5).toString('ascii')!=='%PDF-')return err('Файл не похож на PDF-документ');let extracted='';try{const parsed=await pdfParse(buf);extracted=parsed.text||'';}catch{}const found=extracted.match(/BYPR\d{10}/i)?.[0]||'';let certId=null;if(found){let c=(await pool.query('SELECT id FROM certificates WHERE number=$1',[found])).rows[0];if(!c){const mans=(await pool.query('SELECT id,name FROM manufacturers')).rows;const man=mans.find(m=>extracted.toLowerCase().includes(String(m.name).toLowerCase()));const made=await pool.query(`INSERT INTO certificates(number,manufacturer_id,status,notes) VALUES($1,$2,'На проверке','Номер извлечён из загруженного PDF; реквизиты и срок действия необходимо сверить с оригиналом') RETURNING id`,[found,man?.id||null]);c=made.rows[0];}certId=c.id;}const r=await pool.query(`INSERT INTO certificate_files(certificate_id,original_name,stored_name,extracted_text,file_data) VALUES($1,$2,$3,$4,$5) RETURNING id`,[certId,file.name,crypto.randomUUID()+'.pdf',extracted,buf]);return json({id:r.rows[0].id,number_detected:found||null});
    }
    if(path==='/api/import/products'&&method==='POST'){
      const fd=await req.formData(),file=fd.get('file');if(!file||!file.name?.toLowerCase().endsWith('.xlsx'))return err('Загрузите файл XLSX');if(file.size>5*1024*1024)return err('Максимальный размер Excel — 5 МБ');const wb=XLSX.read(Buffer.from(await file.arrayBuffer()),{type:'buffer'});const sheet=wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(sheet,{defval:''});let added=0,updated=0;const c=await pool.connect();try{await c.query('BEGIN');for(const row of rows){const pick=(...keys)=>{for(const k of keys)if(row[k]!==undefined)return row[k];return ''};const article=text(pick('Артикул','article','Артикул товара'));const name=text(pick('Наименование','Наименование товара','Название','name'));if(!article||!name)continue;const mn=text(pick('Производитель','manufacturer'));let mid=null;if(mn){const mr=await c.query('INSERT INTO manufacturers(name) VALUES($1) ON CONFLICT(name) DO UPDATE SET name=EXCLUDED.name RETURNING id',[mn]);mid=mr.rows[0].id;}const vals=[article,name,text(pick('Категория','category')),text(pick('ТН ВЭД','ТНВЭД','Код ТН ВЭД','Код ТН ВЭД ЕАЭС','tnved')),mid,text(pick('Единица','Ед.изм.','unit'),'шт'),num(pick('НДС','vat'),20),num(pick('Цена закупки','Закупочная цена','purchase_price')),num(pick('Цена продажи','Цена','sale_price'))];const exists=(await c.query('SELECT id FROM products WHERE article=$1',[article])).rows[0];await c.query(`INSERT INTO products(article,name,category,tnved,manufacturer_id,unit,vat,purchase_price,sale_price,sale_price_without_vat) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10) ON CONFLICT(article) DO UPDATE SET name=EXCLUDED.name,category=EXCLUDED.category,tnved=EXCLUDED.tnved,manufacturer_id=EXCLUDED.manufacturer_id,unit=EXCLUDED.unit,vat=EXCLUDED.vat,purchase_price=EXCLUDED.purchase_price,sale_price=EXCLUDED.sale_price,sale_price_without_vat=EXCLUDED.sale_price_without_vat,active=TRUE`,[...vals,vals[8]/(1+vals[6]/100)]);if(exists)updated++;else added++;}await c.query('COMMIT');return json({added,updated});}catch(e){await c.query('ROLLBACK');throw e}finally{c.release()}
    }
    if(path==='/api/import/certificate-list')return err('Список продукции из исходного DOC уже включён в базу (241 строка). Повторный импорт старого формата DOC в веб-версии отключён; обновление списка выполняется отдельной подготовленной загрузкой.',400);
    const d=await body(req);
    if(path==='/api/products'&&method==='POST'){const r=await pool.query(`INSERT INTO products(article,name,category,tnved,manufacturer_id,unit,vat,purchase_price,sale_price,sale_price_without_vat,notes) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) RETURNING id`,[text(d.article),text(d.name),text(d.category),text(d.tnved),d.manufacturer_id||null,text(d.unit,'шт'),num(d.vat,20),num(d.purchase_price),num(d.sale_price),num(d.sale_price_without_vat),text(d.notes)]);return json({id:r.rows[0].id});}
    const productMatch=path.match(/^\/api\/products\/(\d+)$/);if(productMatch){const id=Number(productMatch[1]);if(method==='PATCH'){if(d.restore===true){const r=await pool.query('UPDATE products SET active=TRUE WHERE id=$1 AND active=FALSE RETURNING id',[id]);if(!r.rowCount)return err('Товар не найден в архиве',404);return json({ok:true,restored:true});}await pool.query('UPDATE products SET active=FALSE WHERE id=$1',[id]);return json({ok:true});}await pool.query(`UPDATE products SET article=$1,name=$2,category=$3,tnved=$4,manufacturer_id=$5,unit=$6,vat=$7,purchase_price=$8,sale_price=$9,sale_price_without_vat=$10,notes=$11 WHERE id=$12`,[text(d.article),text(d.name),text(d.category),text(d.tnved),d.manufacturer_id||null,text(d.unit,'шт'),num(d.vat,20),num(d.purchase_price),num(d.sale_price),num(d.sale_price_without_vat),text(d.notes),id]);return json({ok:true});}
    const simpleMatch=path.match(/^\/api\/(customers|suppliers|manufacturers)\/(\d+)$/);
    if(simpleMatch&&method==='PUT'){
      const kind=simpleMatch[1],id=Number(simpleMatch[2]);
      if(kind==='customers') await pool.query('UPDATE customers SET name=$1,unp=$2,address=$3,phone=$4,email=$5,bank_details=$6,notes=$7 WHERE id=$8',[text(d.name),text(d.unp),text(d.address),text(d.phone),text(d.email),text(d.bank_details),text(d.notes),id]);
      if(kind==='suppliers') await pool.query('UPDATE suppliers SET name=$1,unp=$2,phone=$3,email=$4,address=$5,notes=$6 WHERE id=$7',[text(d.name),text(d.unp),text(d.phone),text(d.email),text(d.address),text(d.notes),id]);
      if(kind==='manufacturers') await pool.query('UPDATE manufacturers SET name=$1,country=$2,details=$3 WHERE id=$4',[text(d.name),text(d.country),text(d.details),id]);
      return json({ok:true});
    }
    if(path==='/api/customers'){const r=await pool.query(`INSERT INTO customers(name,unp,address,phone,email,bank_details,notes) VALUES($1,$2,$3,$4,$5,$6,$7) RETURNING id`,[text(d.name),text(d.unp),text(d.address),text(d.phone),text(d.email),text(d.bank_details),text(d.notes)]);return json({id:r.rows[0].id});}
    if(path==='/api/suppliers'){const r=await pool.query(`INSERT INTO suppliers(name,unp,phone,email,address,notes) VALUES($1,$2,$3,$4,$5,$6) RETURNING id`,[text(d.name),text(d.unp),text(d.phone),text(d.email),text(d.address),text(d.notes)]);return json({id:r.rows[0].id});}
    if(path==='/api/manufacturers'){const r=await pool.query(`INSERT INTO manufacturers(name,country,details) VALUES($1,$2,$3) RETURNING id`,[text(d.name),text(d.country),text(d.details)]);return json({id:r.rows[0].id});}
    if(path==='/api/certificates'&&method==='POST'){const r=await pool.query(`INSERT INTO certificates(number,doc_type,manufacturer_id,issue_date,expiry_date,status,notes) VALUES($1,$2,$3,$4,$5,$6,$7) RETURNING id`,[text(d.number),text(d.doc_type,'Сертификат'),d.manufacturer_id||null,text(d.issue_date),text(d.expiry_date),text(d.status,'На проверке'),text(d.notes)]);return json({id:r.rows[0].id});}
    const certMatch=path.match(/^\/api\/certificates\/(\d+)$/);if(certMatch&&method==='PUT'){const id=Number(certMatch[1]);const old=(await pool.query('SELECT number FROM certificates WHERE id=$1',[id])).rows[0];if(!old)return err('Документ не найден',404);await pool.query(`UPDATE certificates SET number=$1,doc_type=$2,manufacturer_id=$3,issue_date=$4,expiry_date=$5,status=$6,notes=$7 WHERE id=$8`,[text(d.number),text(d.doc_type,'Сертификат'),d.manufacturer_id||null,text(d.issue_date),text(d.expiry_date),text(d.status,'На проверке'),text(d.notes),id]);if(old.number!==text(d.number))await pool.query('UPDATE certificate_catalog_items SET certificate_number=$1 WHERE certificate_number=$2',[text(d.number),old.number]);return json({ok:true});}
    const warehouseMatch=path.match(/^\/api\/warehouses\/(\d+)$/);
    if(warehouseMatch&&method==='PUT'){await pool.query('UPDATE warehouses SET name=$1,address=$2 WHERE id=$3',[text(d.name),text(d.address),Number(warehouseMatch[1])]);return json({ok:true});}
    if(path==='/api/warehouses'){const r=await pool.query('INSERT INTO warehouses(name,address) VALUES($1,$2) RETURNING id',[text(d.name),text(d.address)]);return json({id:r.rows[0].id});}
    if(path==='/api/product-certificates'){await pool.query('INSERT INTO product_certificates(product_id,certificate_id) VALUES($1,$2) ON CONFLICT DO NOTHING',[Number(d.product_id),Number(d.certificate_id)]);return json({ok:true});}
    const link=path.match(/^\/api\/certificate-catalog\/(\d+)\/link$/);if(link){const id=Number(link[1]),pid=Number(d.product_id);const c=await pool.connect();try{await c.query('BEGIN');const item=(await c.query('SELECT * FROM certificate_catalog_items WHERE id=$1 FOR UPDATE',[id])).rows[0];if(!item)throw new Error('Строка каталога не найдена');const prod=(await c.query('SELECT * FROM products WHERE id=$1 AND active=TRUE',[pid])).rows[0];if(!prod)throw new Error('Активный товар не найден');const cert=(await c.query('SELECT id FROM certificates WHERE number=$1',[item.certificate_number])).rows[0];if(!cert)throw new Error('Сертификат отсутствует в реестре');await c.query(`UPDATE certificate_catalog_items SET product_id=$1,review_status='Сопоставлено — проверить оригинал' WHERE id=$2`,[pid,id]);await c.query('INSERT INTO product_certificates(product_id,certificate_id) VALUES($1,$2) ON CONFLICT DO NOTHING',[pid,cert.id]);if(!text(prod.tnved)&&item.tnved)await c.query('UPDATE products SET tnved=$1 WHERE id=$2',[item.tnved,pid]);await audit(c,'Сопоставление сертификата','certificate_catalog_item',id,`Товар ID ${pid}; сертификат ${item.certificate_number}`);await c.query('COMMIT');return json({ok:true,product_id:pid,certificate_id:cert.id});}catch(e){await c.query('ROLLBACK');return err(e.message)}finally{c.release()}}
    if(path==='/api/stock/receive'||path==='/api/stock/return'){const qty=num(d.qty),price=num(d.purchase_price);if(qty<=0||price<0)return err('Количество должно быть больше нуля, закупочная цена — неотрицательной');const isReturn=path.endsWith('/return');const c=await pool.connect();try{await c.query('BEGIN');const locked=(await c.query('SELECT id FROM products WHERE id=$1 FOR UPDATE',[Number(d.product_id)])).rows[0];if(!locked)throw new Error('Товар не найден');const r=await c.query(`INSERT INTO batches(product_id,warehouse_id,supplier_id,batch_code,received_at,qty_initial,qty_remaining,purchase_price,certificate_id,notes) VALUES($1,$2,$3,$4,$5,$6,$6,$7,$8,$9) RETURNING id`,[Number(d.product_id),Number(d.warehouse_id),isReturn?null:(d.supplier_id||null),isReturn?'ВОЗВРАТ '+text(d.ref):text(d.batch_code),text(d.received_at,today()),qty,price,d.certificate_id||null,text(d.notes,isReturn?'Возврат от покупателя':'')]);await c.query(`INSERT INTO movements(kind,product_id,warehouse_id,batch_id,qty,unit_price,ref,party,notes) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9)`,[isReturn?'Возврат от покупателя':'Приход',Number(d.product_id),Number(d.warehouse_id),r.rows[0].id,qty,price,text(d.ref||d.batch_code),text(d.party),text(d.notes)]);await c.query('COMMIT');return json({id:r.rows[0].id});}catch(e){await c.query('ROLLBACK');throw e}finally{c.release()}}
    if(path==='/api/stock/issue'){const qty=num(d.qty);const c=await pool.connect();try{await c.query('BEGIN');const used=await fifoDeduct(c,Number(d.product_id),Number(d.warehouse_id),qty,text(d.ref),text(d.party));await c.query('COMMIT');return json({ok:true,batches:used});}catch(e){await c.query('ROLLBACK');return err(e.message)}finally{c.release()}}
    if(path==='/api/stock/transfer'){const qty=num(d.qty);if(qty<=0||Number(d.from_warehouse_id)===Number(d.to_warehouse_id))return err('Проверьте количество и склады');const c=await pool.connect();try{await c.query('BEGIN');const used=await fifoDeduct(c,Number(d.product_id),Number(d.from_warehouse_id),qty,'Перемещение',text(d.ref),'Перемещение','Перемещение — расход');for(const u of used){const b=(await c.query('SELECT * FROM batches WHERE id=$1',[u.batch_id])).rows[0];const ins=await c.query(`INSERT INTO batches(product_id,warehouse_id,batch_code,received_at,qty_initial,qty_remaining,purchase_price,certificate_id,notes) VALUES($1,$2,$3,$4,$5,$5,$6,$7,$8) RETURNING id`,[b.product_id,Number(d.to_warehouse_id),`Перемещение из партии ${b.id}`,today(),u.qty,b.purchase_price,b.certificate_id,text(d.notes)]);await c.query(`INSERT INTO movements(kind,product_id,warehouse_id,batch_id,qty,unit_price,ref,notes) VALUES('Перемещение — приход',$1,$2,$3,$4,$5,$6,$7)`,[b.product_id,Number(d.to_warehouse_id),ins.rows[0].id,u.qty,b.purchase_price,text(d.ref),text(d.notes)]);}await c.query('COMMIT');return json({ok:true});}catch(e){await c.query('ROLLBACK');return err(e.message)}finally{c.release()}}
    if(path==='/api/stock/adjust'){const pid=Number(d.product_id),wid=Number(d.warehouse_id),actual=num(d.actual_qty);if(actual<0)return err('Фактический остаток не может быть отрицательным');const c=await pool.connect();try{await c.query('BEGIN');await c.query('SELECT id FROM products WHERE id=$1 FOR UPDATE',[pid]);await c.query('SELECT id FROM batches WHERE product_id=$1 AND warehouse_id=$2 FOR UPDATE',[pid,wid]);const r=await c.query('SELECT COALESCE(SUM(qty_remaining),0)::float qty FROM batches WHERE product_id=$1 AND warehouse_id=$2',[pid,wid]);const diff=actual-Number(r.rows[0].qty);if(diff>0){const b=await c.query(`INSERT INTO batches(product_id,warehouse_id,batch_code,received_at,qty_initial,qty_remaining,notes) VALUES($1,$2,'ИНВЕНТАРИЗАЦИЯ',$3,$4,$4,$5) RETURNING id`,[pid,wid,today(),diff,text(d.notes)]);await c.query(`INSERT INTO movements(kind,product_id,warehouse_id,batch_id,qty,ref,notes) VALUES('Инвентаризация',$1,$2,$3,$4,'Пересчёт',$5)`,[pid,wid,b.rows[0].id,diff,text(d.notes)]);}else if(diff<0){await fifoDeduct(c,pid,wid,-diff,'Инвентаризация','',text(d.notes,'Корректировка'));}await c.query('COMMIT');return json({ok:true,difference:diff});}catch(e){await c.query('ROLLBACK');return err(e.message)}finally{c.release()}}
    if(path==='/api/documents'&&method==='POST'){const items=Array.isArray(d.items)?d.items:[];if(!items.length)return err('Добавьте хотя бы одну товарную позицию');const c=await pool.connect();try{await c.query('BEGIN');for(const it of items){if(num(it.qty)<=0||num(it.price,-1)<0)throw new Error('Количество должно быть больше нуля, цена — неотрицательной');const p=(await c.query('SELECT id FROM products WHERE id=$1 AND active=TRUE',[Number(it.product_id)])).rows[0];if(!p)throw new Error('В документе указан неизвестный или архивный товар');}const r=await c.query(`INSERT INTO documents(doc_no,doc_date,customer_id,status,doc_type,notes,warehouse_id,vat_mode) VALUES($1,$2,$3,'Черновик',$4,$5,$6,$7) RETURNING id`,[text(d.doc_no),text(d.doc_date,today()),d.customer_id||null,text(d.doc_type,'Счёт'),text(d.notes),d.warehouse_id||null,['inclusive','exclusive','none'].includes(d.vat_mode)?d.vat_mode:'inclusive']);const id=r.rows[0].id;for(const it of items){const p=(await c.query(`SELECT p.*,m.name manufacturer,(SELECT string_agg(ce.number,', ') FROM product_certificates pc JOIN certificates ce ON ce.id=pc.certificate_id WHERE pc.product_id=p.id) certificates FROM products p LEFT JOIN manufacturers m ON m.id=p.manufacturer_id WHERE p.id=$1`,[Number(it.product_id)])).rows[0];await c.query(`INSERT INTO document_items(document_id,product_id,article,name,tnved,manufacturer,certificate,qty,price,vat) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)`,[id,p.id,p.article,p.name,p.tnved,p.manufacturer||'',p.certificates||'',num(it.qty),num(it.price),num(p.vat)]);}await c.query('COMMIT');return json({id});}catch(e){await c.query('ROLLBACK');return err(e.message)}finally{c.release()}}
    const post=path.match(/^\/api\/documents\/(\d+)\/post$/);if(post){const id=Number(post[1]),c=await pool.connect();try{await c.query('BEGIN');const doc=(await c.query('SELECT * FROM documents WHERE id=$1 FOR UPDATE',[id])).rows[0];if(!doc)throw new Error('Документ не найден');if(doc.status!=='Черновик')throw new Error('Провести можно только черновик');if(!doc.warehouse_id)throw new Error('В документе не выбран склад');const its=(await c.query('SELECT * FROM document_items WHERE document_id=$1 ORDER BY product_id,id',[id])).rows;for(const it of its){const used=await fifoDeduct(c,it.product_id,doc.warehouse_id,Number(it.qty),doc.doc_no,'',`Реализация ${doc.doc_no}`,'Реализация');for(const u of used)await c.query('INSERT INTO document_allocations(document_item_id,batch_id,qty) VALUES($1,$2,$3)',[it.id,u.batch_id,u.qty]);}await c.query(`UPDATE documents SET status='Проведён' WHERE id=$1`,[id]);await audit(c,'Проведение документа','document',id,doc.doc_no);await c.query('COMMIT');return json({ok:true});}catch(e){await c.query('ROLLBACK');return err(e.message)}finally{c.release()}}
    const cancel=path.match(/^\/api\/documents\/(\d+)\/cancel$/);if(cancel){const id=Number(cancel[1]),c=await pool.connect();try{await c.query('BEGIN');const doc=(await c.query('SELECT * FROM documents WHERE id=$1 FOR UPDATE',[id])).rows[0];if(!doc||doc.status!=='Проведён')throw new Error('Документ не найден или не проведён');const alloc=(await c.query(`SELECT da.*,di.product_id FROM document_allocations da JOIN document_items di ON di.id=da.document_item_id WHERE di.document_id=$1`,[id])).rows;for(const a of alloc){await c.query('UPDATE batches SET qty_remaining=qty_remaining+$1 WHERE id=$2',[a.qty,a.batch_id]);await c.query(`INSERT INTO movements(kind,product_id,warehouse_id,batch_id,qty,ref,notes) VALUES('Отмена реализации',$1,$2,$3,$4,$5,$6)`,[a.product_id,doc.warehouse_id,a.batch_id,a.qty,doc.doc_no,'Возврат остатка после отмены']);}await c.query('DELETE FROM document_allocations WHERE document_item_id IN (SELECT id FROM document_items WHERE document_id=$1)',[id]);await c.query(`UPDATE documents SET status='Отменён' WHERE id=$1`,[id]);await audit(c,'Отмена документа','document',id,doc.doc_no);await c.query('COMMIT');return json({ok:true});}catch(e){await c.query('ROLLBACK');return err(e.message)}finally{c.release()}}
  }
  if(method==='DELETE'){
    const m=path.match(/^\/api\/(products|customers|suppliers|manufacturers|warehouses|certificates)\/(\d+)$/);
    if(!m){
      const dm=path.match(/^\/api\/documents\/(\d+)$/);
      if(dm){
        const id=Number(dm[1]),doc=(await pool.query('SELECT status FROM documents WHERE id=$1',[id])).rows[0];
        if(!doc)return err('Документ не найден',404);
        if(doc.status==='Проведён')return err('Сначала отмените проведение документа. После этого его можно удалить.',409);
        await pool.query('DELETE FROM documents WHERE id=$1',[id]);return json({ok:true});
      }
      return err('Неизвестный маршрут',404);
    }
    const kind=m[1],id=Number(m[2]),c=await pool.connect();
    try{
      await c.query('BEGIN');
      const checks={
        products:[['batches','product_id'],['movements','product_id'],['document_items','product_id'],['product_certificates','product_id'],['certificate_catalog_items','product_id']],
        customers:[['documents','customer_id']],
        suppliers:[['batches','supplier_id']],
        manufacturers:[['products','manufacturer_id'],['certificates','manufacturer_id']],
        warehouses:[['batches','warehouse_id'],['movements','warehouse_id'],['documents','warehouse_id']],
        certificates:[['batches','certificate_id'],['product_certificates','certificate_id'],['certificate_files','certificate_id']]
      };
      if(kind==='products'){
        const used=await c.query(`SELECT
          EXISTS(SELECT 1 FROM batches WHERE product_id=$1) OR
          EXISTS(SELECT 1 FROM movements WHERE product_id=$1) OR
          EXISTS(SELECT 1 FROM document_items WHERE product_id=$1) AS used`,[id]);
        if(used.rows[0].used){await c.query('UPDATE products SET active=FALSE WHERE id=$1',[id]);await c.query('COMMIT');return json({ok:true,archived:true,message:'У товара есть история или складские операции — он перемещён в архив'});}
      }
      if(kind==='certificates'){
        const cert=(await c.query('SELECT number FROM certificates WHERE id=$1',[id])).rows[0];
        if(cert && (await c.query('SELECT 1 FROM certificate_catalog_items WHERE certificate_number=$1 LIMIT 1',[cert.number])).rowCount) throw new Error('Нельзя удалить: с этим номером связаны позиции каталога сертификатов.');
      }
      for(const [table,col] of checks[kind]){
        const q=await c.query(`SELECT 1 FROM ${table} WHERE ${col}=$1 LIMIT 1`,[id]);
        if(q.rowCount)throw new Error('Нельзя удалить: запись используется в истории или связанных данных. Сначала проверьте связанные операции.');
      }
      const table=kind;
      await c.query(`DELETE FROM ${table} WHERE id=$1`,[id]);
      await c.query('COMMIT');return json({ok:true});
    }catch(e){await c.query('ROLLBACK');return err(e.message,409)}finally{c.release()}
  }
  return err('Неизвестный маршрут',404);
}
export default async (req) => {
  try { return await handler(req); }
  catch (e) { console.error('API error',e); return json({error:'Внутренняя ошибка сервера. Проверьте журнал Netlify.'},500); }
};
