const SPREADSHEET_ID = '1rlbfQNIlt2X3F8My_auY0MPrSYhmFdmIbhaaKs6gPFs';

function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const schemas = {
    Products: ['ID','Name','Description','Image','Category','OriginalPrice','SalePrice','Stock','Active','CreatedAt'],
    Orders: ['OrderID','CustomerName','Phone','Items','Quantity','Total','Notes','Status','CreatedAt'],
    Settings: ['Key','Value']
  };
  Object.entries(schemas).forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) { sh.appendRow(headers); sh.setFrozenRows(1); }
  });
  return 'NEXA STORE DATABASE READY';
}
function json(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function getSheet(name) { const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name); if(!sh) throw new Error('Missing sheet: '+name); return sh; }
function getAdminKey() { return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY'); }
function requireAdmin(data) { const saved=getAdminKey(); if(!saved || !data.adminKey || data.adminKey!==saved) throw new Error('بيانات دخول الإدارة غير صحيحة'); }
function rowsAsObjects(sheet) { const values=sheet.getDataRange().getValues(); if(values.length<2)return []; const headers=values.shift(); return values.map(row=>{const o={};headers.forEach((h,i)=>o[h]=row[i]);return o;}); }
function doGet(e) {
  const action=(e && e.parameter && e.parameter.action)||'health';
  if(action==='health') return json({success:true,store:'NEXA STORE',status:'Backend connected'});
  if(action==='products') {
    const products=rowsAsObjects(getSheet('Products')).map(p=>({
      id:String(p.ID),name:String(p.Name||''),description:String(p.Description||''),image:String(p.Image||''),category:String(p.Category||''),
      originalPrice:p.OriginalPrice===''?'':Number(p.OriginalPrice||0),price:p.SalePrice===''?'':Number(p.SalePrice||0),
      stock:Number(p.Stock||0),active:p.Active===true||String(p.Active).toLowerCase()==='true'
    })).filter(p=>p.id && p.active);
    return json({success:true,products:products});
  }
  return json({success:false,error:'Unknown action'});
}
function doPost(e) {
  try {
    const data=JSON.parse((e && e.postData && e.postData.contents)||'{}');
    if(data.action==='createOrder') return createOrder(data);
    requireAdmin(data);
    if(data.action==='saveProduct') return saveProduct(data);
    if(data.action==='deleteProduct') return deleteProduct(data);
    if(data.action==='updateOrder') return updateOrder(data);
    if(data.action==='getOrders') return getOrders();
    return json({success:false,error:'Unknown action'});
  } catch(err) { return json({success:false,error:err.message||'Request failed'}); }
}
function saveProduct(data) {
  const p=data.product||{}; if(!String(p.name||'').trim()) throw new Error('اسم المنتج مطلوب');
  const price=Number(p.price||0), original=Number(p.originalPrice||0), stock=Number(p.stock||0);
  if(!Number.isFinite(price)||price<0||!Number.isFinite(original)||original<0||!Number.isFinite(stock)||stock<0) throw new Error('تحقق من الأسعار والمخزون');
  const sh=getSheet('Products'), values=sh.getDataRange().getValues(), id=String(p.id||Utilities.getUuid());
  const row=[id,String(p.name).trim(),String(p.description||''),String(p.image||''),String(p.category||''),original,price,stock,p.active!==false,new Date()];
  const index=values.findIndex((r,i)=>i>0&&String(r[0])===id);
  if(index>0) sh.getRange(index+1,1,1,row.length).setValues([row]); else sh.appendRow(row);
  return json({success:true,id:id});
}
function deleteProduct(data) {
  if(!data.id) throw new Error('Product ID required'); const sh=getSheet('Products'), values=sh.getDataRange().getValues();
  for(let i=1;i<values.length;i++) if(String(values[i][0])===String(data.id)){sh.deleteRow(i+1);return json({success:true});}
  throw new Error('Product not found');
}
function createOrder(data) {
  const name=String(data.customerName||'').trim(), phone=String(data.phone||'').trim(), items=data.items;
  if(!name||!phone||!Array.isArray(items)||!items.length) throw new Error('الاسم والهاتف والمنتجات مطلوبة');
  const sh=getSheet('Products'), products=rowsAsObjects(sh), lookup={}; products.forEach(p=>lookup[String(p.ID)]=p);
  let total=0, quantity=0; const clean=[];
  items.forEach(item=>{const p=lookup[String(item.id)];const q=Math.floor(Number(item.quantity));if(!p||String(p.Active).toLowerCase()==='false'||q<1)throw new Error('منتج غير صالح في الطلب');const price=Number(p.SalePrice||0);total+=price*q;quantity+=q;clean.push({id:String(p.ID),name:String(p.Name),price:price,quantity:q});});
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try { const orderId='NX-'+Utilities.getUuid().slice(0,8).toUpperCase();getSheet('Orders').appendRow([orderId,name,phone,JSON.stringify(clean),quantity,total,String(data.notes||''),'جديد',new Date()]);return json({success:true,orderId:orderId,total:total});}
  finally {lock.releaseLock();}
}
function getOrders() {
  const orders=rowsAsObjects(getSheet('Orders')).reverse().map(o=>({
    orderId:String(o.OrderID),customerName:String(o.CustomerName||''),phone:String(o.Phone||''),items:String(o.Items||''),
    quantity:Number(o.Quantity||0),total:Number(o.Total||0),notes:String(o.Notes||''),status:String(o.Status||'جديد'),
    createdAt:o.CreatedAt instanceof Date?Utilities.formatDate(o.CreatedAt,Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm'):String(o.CreatedAt||'')
  }));
  return json({success:true,orders:orders});
}
function updateOrder(data) {
  if(!data.orderId||!data.status) throw new Error('Order ID and status are required');
  const allowed=['جديد','قيد التجهيز','مكتمل','ملغي'];if(!allowed.includes(String(data.status)))throw new Error('حالة طلب غير صالحة');
  const sh=getSheet('Orders'),values=sh.getDataRange().getValues();
  for(let i=1;i<values.length;i++)if(String(values[i][0])===String(data.orderId)){sh.getRange(i+1,8).setValue(String(data.status));return json({success:true});}
  throw new Error('Order not found');
}