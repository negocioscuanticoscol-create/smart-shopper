// ═══════════════════════════════════════════════════════════
// SMART SHOPPER — Apps Script v3.1 (con CORS)
// ═══════════════════════════════════════════════════════════

const SH = {
  PEDIDOS:'Pedidos', CLIENTES:'Clientes', SUSCRIPCIONES:'Suscripciones',
  VENTAS:'Ventas', ANALISIS:'Análisis de Productos', CATALOGO:'Catálogo'
};
const PC = {
  FECHA:0,CODIGO:1,NOMBRE:2,CELULAR:3,LOCALIDAD:4,BARRIO:5,
  DIRECCION:6,APTO:7,REFERENCIA:8,GPS:9,HORARIO:10,
  PRODUCTOS:11,NOTA:12,ITEMS:13,ESTADO:14,VALOR:15,YAPA:16,DOM_REST:17,COMISION:18
};
const COLORES = {
  'Lista de compra':'#ffffff','Por consignación':'#fff9c4',
  'Pagado':'#c8e6c9','Empacado':'#b3e5fc','Entregado':'#e1bee7'
};

// ── CORS helper ────────────────────────────────
function corsOut(obj) {
  const out = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return out;
}
function corsTxt(t) {
  return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT);
}

// ── doOptions — responde preflight CORS ────────
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── doPost ─────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    inicializarHojas(ss);
    switch(data.accion) {
      case 'actualizar':       actualizarPedido(ss, data);          break;
      case 'suscripcion':      guardarSuscripcion(ss, data);        break;
      case 'comision':         registrarComision(ss, data);         break;
      case 'agregarProducto':  agregarProductoAlCatalogo(ss, data); break;
      default:
        escribirPedido(ss, data);
        actualizarCliente(ss, data);
        actualizarAnalisis(ss, data);
    }
    return corsOut({ ok: true });
  } catch(err) {
    return corsOut({ ok: false, error: err.message });
  }
}

// ── doGet ──────────────────────────────────────
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  if (!p.accion) return corsTxt('Smart Shopper v3.1 OK');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  inicializarHojas(ss);
  switch(p.accion) {
    case 'pedidos':       return corsOut({ pedidos:       getPedidos(ss, p.estado) });
    case 'clientes':      return corsOut({ clientes:      getClientes(ss) });
    case 'suscripciones': return corsOut({ suscripciones: getSuscripciones(ss) });
    case 'ventas':        return corsOut({ ventas:        getVentas(ss) });
    case 'analisis':      return corsOut({ analisis:      getAnalisis(ss) });
    case 'catalogo':      return corsOut({ catalogo:      getCatalogo(ss) });
    case 'verificar':     return corsOut(verificarCliente(ss, p.celular));
    // ← Recibir pedido y suscripción por GET (evita CORS)
    case 'nuevoPedido':
      escribirPedido(ss, p);
      actualizarCliente(ss, p);
      actualizarAnalisis(ss, p);
      return corsOut({ ok: true });
    case 'nuevaSuscripcion':
      guardarSuscripcion(ss, p);
      return corsOut({ ok: true });
    case 'actualizarPedido':
      actualizarPedido(ss, p);
      return corsOut({ ok: true });
    default: return corsOut({ error: 'Acción desconocida' });
  }
}

// ── INICIALIZAR HOJAS ──────────────────────────
function inicializarHojas(ss) {
  crearPedidos(ss); crearClientes(ss); crearSuscripciones(ss);
  crearVentas(ss); crearAnalisis(ss); crearCatalogo(ss);
}

// ── PEDIDOS ────────────────────────────────────
function crearPedidos(ss) {
  if (ss.getSheetByName(SH.PEDIDOS)) return;
  const s = ss.insertSheet(SH.PEDIDOS);
  const h = ['Fecha','Código','Nombre','Celular','Localidad','Barrio','Dirección',
    'Apto','Referencia','GPS','Horario','Productos','Nota','Items',
    'Estado','Valor ($)','Yapa','Dom.Rest.','Comisión Plaza'];
  encabezado(s, h, '#1b4332');
  s.setFrozenRows(1); s.setColumnWidth(12, 320);
}
function escribirPedido(ss, d) {
  const s = ss.getSheetByName(SH.PEDIDOS);
  const fecha = fmt(new Date());
  s.appendRow([fecha, d.codigo||'', d.nombre||'', d.celular||'',
    d.localidad||'', d.barrio||'', d.direccion||'', d.apto||'',
    d.referencia||'', d.gps||'', d.horario||'', d.productos||'',
    d.nota||'', d.total_items||0, 'Lista de compra', '', '', d.domicilios_restantes||0, '']);
  const lr = s.getLastRow();
  if(lr%2===0) s.getRange(lr,1,1,19).setBackground('#f0faf2');
}
function actualizarPedido(ss, d) {
  const s = ss.getSheetByName(SH.PEDIDOS); if(!s) return;
  const rows = s.getDataRange().getValues();
  for(let i=1; i<rows.length; i++) {
    if(String(rows[i][PC.CODIGO])===String(d.codigo)) {
      const f=i+1;
      if(d.estado!==undefined)   s.getRange(f, PC.ESTADO+1).setValue(d.estado);
      if(d.valor!==undefined)    s.getRange(f, PC.VALOR+1).setValue(d.valor);
      if(d.yapa!==undefined)     s.getRange(f, PC.YAPA+1).setValue(d.yapa);
      if(d.comision!==undefined) s.getRange(f, PC.COMISION+1).setValue(d.comision);
      const color = COLORES[d.estado]||'#ffffff';
      s.getRange(f,1,1,19).setBackground(color);
      if(d.estado==='Entregado' && d.valor) registrarVentaDesde(ss, rows[i], d);
      break;
    }
  }
  if(d.domicilios!==undefined && d.celular) actualizarDomCliente(ss, d.celular, d.domicilios);
}
function getPedidos(ss, estado) {
  const s = ss.getSheetByName(SH.PEDIDOS); if(!s) return [];
  return ss.getSheetByName(SH.PEDIDOS).getDataRange().getValues().slice(1).map((r,i)=>({
    fila:i+2, fecha:String(r[PC.FECHA]), codigo:r[PC.CODIGO],
    nombre:r[PC.NOMBRE], celular:r[PC.CELULAR], localidad:r[PC.LOCALIDAD],
    barrio:r[PC.BARRIO], direccion:r[PC.DIRECCION], apto:r[PC.APTO],
    referencia:r[PC.REFERENCIA], gps:r[PC.GPS], horario:r[PC.HORARIO],
    productos:r[PC.PRODUCTOS], nota:r[PC.NOTA], items:r[PC.ITEMS],
    estado:r[PC.ESTADO]||'Lista de compra', valor:r[PC.VALOR]||'',
    yapa:r[PC.YAPA]||'', domRestantes:r[PC.DOM_REST]||0, comision:r[PC.COMISION]||''
  })).filter(p=>!estado||p.estado===estado);
}

// ── Verificar domicilios de un celular específico ──
function verificarCliente(ss, celular) {
  if(!celular) return {domicilios:0,encontrado:false};
  const cel=String(celular).replace(/\D/g,'');
  // Buscar en Clientes
  const cs=ss.getSheetByName(SH.CLIENTES);
  if(cs){
    const rows=cs.getDataRange().getValues();
    for(let i=1;i<rows.length;i++){
      if(String(rows[i][0]).replace(/\D/g,'')===cel){
        return{domicilios:Number(rows[i][6])||0,encontrado:true,nombre:rows[i][1]};
      }
    }
  }
  // Si no está en Clientes, buscar en Suscripciones
  const ss2=ss.getSheetByName(SH.SUSCRIPCIONES);
  if(ss2){
    const rows=ss2.getDataRange().getValues();
    for(let i=1;i<rows.length;i++){
      if(String(rows[i][0]).replace(/\D/g,'')===cel){
        return{domicilios:Number(rows[i][7])||0,encontrado:true,nombre:rows[i][1]};
      }
    }
  }
  return{domicilios:0,encontrado:false};
}

// ── CLIENTES ───────────────────────────────────
function crearClientes(ss) {
  if(ss.getSheetByName(SH.CLIENTES)) return;
  const s = ss.insertSheet(SH.CLIENTES);
  encabezado(s,['Celular','Nombre','Dirección','Primer Pedido','Último Pedido','Total Pedidos','Domicilios'],'#2d6a4f');
  s.setFrozenRows(1);
}
function actualizarCliente(ss, d) {
  const s=ss.getSheetByName(SH.CLIENTES); if(!s)return;
  const cel=String(d.celular), rows=s.getDataRange().getValues();
  let f=-1;
  for(let i=1;i<rows.length;i++){if(String(rows[i][0])===cel){f=i+1;break;}}
  const ahora=fmt(new Date());
  if(f===-1){s.appendRow([cel,d.nombre||'',d.direccion||'',ahora,ahora,1,d.domicilios_restantes||0]);}
  else{
    s.getRange(f,2).setValue(d.nombre||'');
    s.getRange(f,3).setValue(d.direccion||'');
    s.getRange(f,5).setValue(ahora);
    s.getRange(f,6).setValue((Number(rows[f-1][5])||0)+1);
  }
}
function actualizarDomCliente(ss,celular,cantidad){
  const s=ss.getSheetByName(SH.CLIENTES);if(!s)return;
  const rows=s.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0])===String(celular)){s.getRange(i+1,7).setValue(Number(cantidad));break;}
  }
}
function getClientes(ss){
  const s=ss.getSheetByName(SH.CLIENTES);if(!s)return[];
  return s.getDataRange().getValues().slice(1).map(r=>({
    celular:r[0],nombre:r[1],direccion:r[2],primerPedido:String(r[3]),
    ultimoPedido:String(r[4]),totalPedidos:r[5],domicilios:r[6]||0
  }));
}

// ── SUSCRIPCIONES ──────────────────────────────
function crearSuscripciones(ss){
  if(ss.getSheetByName(SH.SUSCRIPCIONES))return;
  const s=ss.insertSheet(SH.SUSCRIPCIONES);
  encabezado(s,['Celular','Nombre','Plan','Valor Pagado','Fecha Pago',
    'Domicilios Totales','Domicilios Usados','Domicilios Restantes','Estado','Vence'],'#1565c0');
  s.setFrozenRows(1);
}
function guardarSuscripcion(ss,d){
  const s=ss.getSheetByName(SH.SUSCRIPCIONES);if(!s)return;
  const cel=String(d.celular),rows=s.getDataRange().getValues();
  let f=-1;
  for(let i=1;i<rows.length;i++){if(String(rows[i][0])===cel){f=i+1;break;}}
  const fecha=fmt(new Date());
  const vence=d.meses>0?fmtD(addMeses(new Date(),d.meses)):'N/A';
  const envios=Number(d.envios)||0;
  if(f===-1){
    s.appendRow([cel,d.nombre||'',d.plan||'',d.valor||0,fecha,envios,0,envios,'Activo',vence]);
  } else {
    const domAct=Number(rows[f-1][7])||0;
    const domTot=Number(rows[f-1][5])||0;
    s.getRange(f,4).setValue(d.valor);s.getRange(f,5).setValue(fecha);
    s.getRange(f,6).setValue(domTot+envios);
    s.getRange(f,8).setValue(domAct+envios);
    s.getRange(f,10).setValue(vence);
  }
  // ← CRÍTICO: también crear/actualizar en pestaña Clientes
  crearClientes(ss);
  const cs=ss.getSheetByName(SH.CLIENTES);
  const crow=cs.getDataRange().getValues();
  let cf=-1;
  for(let i=1;i<crow.length;i++){if(String(crow[i][0])===cel){cf=i+1;break;}}
  const domActuales=cf>-1?(Number(crow[cf-1][6])||0):0;
  if(cf===-1){
    cs.appendRow([cel,d.nombre||'','',fecha,fecha,0,envios]);
  } else {
    cs.getRange(cf,7).setValue(domActuales+envios);
  }
}
function getSuscripciones(ss){
  const s=ss.getSheetByName(SH.SUSCRIPCIONES);if(!s)return[];
  return s.getDataRange().getValues().slice(1).map(r=>({
    celular:r[0],nombre:r[1],plan:r[2],valorPagado:r[3],fechaPago:String(r[4]),
    domTotales:r[5],domUsados:r[6],domRestantes:r[7],estado:r[8],vence:String(r[9])
  }));
}

// ── VENTAS ─────────────────────────────────────
function crearVentas(ss){
  if(ss.getSheetByName(SH.VENTAS))return;
  const s=ss.insertSheet(SH.VENTAS);
  encabezado(s,['Fecha','Código','Cliente','Celular','Localidad',
    'Valor ($)','Comisión Plaza ($)','% Plaza','Ingreso Neto ($)','Mes'],'#880e4f');
  s.setFrozenRows(1);
}
function registrarVentaDesde(ss, row, d) {
  const s=ss.getSheetByName(SH.VENTAS);if(!s)return;
  const valor=Number(d.valor)||0, comision=Number(d.comision||row[PC.COMISION])||0;
  const mes=new Date().toLocaleDateString('es-CO',{month:'long',year:'numeric'});
  s.appendRow([fmt(new Date()),row[PC.CODIGO],row[PC.NOMBRE],row[PC.CELULAR],
    row[PC.LOCALIDAD],valor,comision,
    valor>0?Math.round(comision/valor*100)+'%':'',valor-comision,mes]);
}
function registrarComision(ss,d){ actualizarPedido(ss,d); }
function getVentas(ss){
  const s=ss.getSheetByName(SH.VENTAS);if(!s)return{filas:[],totales:{}};
  const filas=s.getDataRange().getValues().slice(1).map(r=>({
    fecha:String(r[0]),codigo:r[1],cliente:r[2],celular:r[3],localidad:r[4],
    valor:r[5],comision:r[6],pctPlaza:r[7],neto:r[8],mes:r[9]
  }));
  return{filas,totales:{
    totalVentas:filas.reduce((a,r)=>a+Number(r.valor||0),0),
    totalComisiones:filas.reduce((a,r)=>a+Number(r.comision||0),0),
    totalNeto:filas.reduce((a,r)=>a+Number(r.neto||0),0),
    numPedidos:filas.length
  }};
}

// ── ANÁLISIS ───────────────────────────────────
function crearAnalisis(ss){
  if(ss.getSheetByName(SH.ANALISIS))return;
  const s=ss.insertSheet(SH.ANALISIS);
  encabezado(s,['Producto','Categoría','Total Pedidos','Última Vez','Este Mes','Mes Anterior','Tendencia','Unidad'],'#1a5c1a');
  s.setFrozenRows(1);
}
function actualizarAnalisis(ss,data){
  const s=ss.getSheetByName(SH.ANALISIS);if(!s||!data.productos)return;
  const prods=data.productos.split('|').map(p=>p.trim()).filter(Boolean);
  const rows=s.getDataRange().getValues();
  const idx={};for(let i=1;i<rows.length;i++)idx[rows[i][0]]=i+1;
  prods.forEach(p=>{
    const nombre=p.split(':')[0].trim();if(!nombre)return;
    if(idx[nombre]){
      const f=idx[nombre],tot=Number(s.getRange(f,3).getValue())||0,em=Number(s.getRange(f,5).getValue())||0;
      s.getRange(f,3).setValue(tot+1);s.getRange(f,4).setValue(fmt(new Date()));s.getRange(f,5).setValue(em+1);
      s.getRange(f,7).setValue(em+1>(Number(s.getRange(f,6).getValue())||0)?'↑ Sube':'↓ Baja');
    }else{s.appendRow([nombre,'',1,fmt(new Date()),1,0,'Nuevo','']);}
  });
  if(s.getLastRow()>2)s.getRange(2,1,s.getLastRow()-1,8).sort({column:3,ascending:false});
}
function getAnalisis(ss){
  const s=ss.getSheetByName(SH.ANALISIS);if(!s)return[];
  return s.getDataRange().getValues().slice(1).map(r=>({
    producto:r[0],categoria:r[1],total:r[2],ultimaVez:String(r[3]),
    esteMes:r[4],mesAnterior:r[5],tendencia:r[6],unidad:r[7]
  }));
}

// ── CATÁLOGO ───────────────────────────────────
function crearCatalogo(ss){
  if(ss.getSheetByName(SH.CATALOGO))return;
  const s=ss.insertSheet(SH.CATALOGO);
  encabezado(s,['Nombre','Categoría','Tipo (u/g)','Activo (SI/NO)'],'#4a148c');
  s.setFrozenRows(1);
  const inicial=[
    ['Acelga','🥦 Verduras','g','SI'],['Auyama','🥦 Verduras','g','SI'],['Alcachofa','🥦 Verduras','u','SI'],
    ['Ají','🥦 Verduras','u','SI'],['Ajo Importado','🥦 Verduras','u','SI'],['Ajo Nacional','🥦 Verduras','u','SI'],
    ['Apio','🥦 Verduras','u','SI'],['Aromáticas','🥦 Verduras','u','SI'],['Arracacha','🥦 Verduras','g','SI'],
    ['Arveja Cáscara','🥦 Verduras','g','SI'],['Arveja Desgranada','🥦 Verduras','g','SI'],
    ['Brócoli','🥦 Verduras','u','SI'],['Berenjena','🥦 Verduras','u','SI'],
    ['Cabezona Blanca','🥦 Verduras','u','SI'],['Cabezona Roja','🥦 Verduras','u','SI'],
    ['Calabaza','🥦 Verduras','g','SI'],['Camote','🥦 Verduras','g','SI'],
    ['Cebolla Larga','🥦 Verduras','u','SI'],['Cebolla Puerro','🥦 Verduras','u','SI'],
    ['Cebollin','🥦 Verduras','u','SI'],['Champiñón','🥦 Verduras','g','SI'],
    ['Cilantro','🥦 Verduras','u','SI'],['Coliflor','🥦 Verduras','u','SI'],
    ['Espárragos','🥦 Verduras','g','SI'],['Espinaca','🥦 Verduras','g','SI'],
    ['Espinaca Baby','🥦 Verduras','g','SI'],['Frijol','🥦 Verduras','g','SI'],
    ['Guatila','🥦 Verduras','u','SI'],['Habichuela','🥦 Verduras','g','SI'],
    ['Jengibre','🥦 Verduras','g','SI'],['Lechuga Cogollo','🥦 Verduras','u','SI'],
    ['Lechuga Batavia','🥦 Verduras','u','SI'],['Lechuga Morada','🥦 Verduras','u','SI'],
    ['Lechuga Romana','🥦 Verduras','u','SI'],['Lechuga Rúgula','🥦 Verduras','g','SI'],
    ['Mazorca','🥦 Verduras','u','SI'],['Mazorca Desgranada','🥦 Verduras','g','SI'],
    ['Papa Criolla','🥦 Verduras','g','SI'],['Papa Sabanera','🥦 Verduras','g','SI'],
    ['Papa Pastusa','🥦 Verduras','g','SI'],['Pepino Cohombro','🥦 Verduras','u','SI'],
    ['Pepino Común','🥦 Verduras','u','SI'],['Perejil','🥦 Verduras','u','SI'],
    ['Pimentón','🥦 Verduras','u','SI'],['Rábano','🥦 Verduras','u','SI'],
    ['Remolacha','🥦 Verduras','g','SI'],['Repollo Blanco','🥦 Verduras','u','SI'],
    ['Repollo Morado','🥦 Verduras','u','SI'],['Tomate Cherry','🥦 Verduras','g','SI'],
    ['Tomate Chonto','🥦 Verduras','g','SI'],['Tomate Milano','🥦 Verduras','g','SI'],
    ['Yuca','🥦 Verduras','g','SI'],['Zanahoria','🥦 Verduras','g','SI'],
    ['Zanahoria Baby','🥦 Verduras','g','SI'],['Zuchini','🥦 Verduras','u','SI'],
    ['Aguacate Lorena','🍎 Frutas','u','SI'],['Aguacate Hass','🍎 Frutas','u','SI'],
    ['Agraz','🍎 Frutas','g','SI'],['Arándanos','🍎 Frutas','g','SI'],
    ['Banano Bocadillo','🍎 Frutas','u','SI'],['Banano Criollo','🍎 Frutas','u','SI'],
    ['Ciruela','🍎 Frutas','g','SI'],['Coco','🍎 Frutas','u','SI'],
    ['Curuba','🍎 Frutas','u','SI'],['Durazno','🍎 Frutas','u','SI'],
    ['Feijoa','🍎 Frutas','u','SI'],['Fresa','🍎 Frutas','g','SI'],
    ['Granadilla','🍎 Frutas','u','SI'],['Guanábana','🍎 Frutas','g','SI'],
    ['Guayaba','🍎 Frutas','g','SI'],['Kiwi','🍎 Frutas','u','SI'],
    ['Limón','🍎 Frutas','u','SI'],['Lulo','🍎 Frutas','u','SI'],
    ['Mandarina','🍎 Frutas','u','SI'],['Mango Azúcar','🍎 Frutas','u','SI'],
    ['Mango Tommy','🍎 Frutas','u','SI'],['Manzana Nacional','🍎 Frutas','u','SI'],
    ['Manzana Roja','🍎 Frutas','u','SI'],['Manzana Verde','🍎 Frutas','u','SI'],
    ['Maracuyá','🍎 Frutas','u','SI'],['Melón','🍎 Frutas','u','SI'],
    ['Mora','🍎 Frutas','g','SI'],['Naranja','🍎 Frutas','u','SI'],
    ['Papaya','🍎 Frutas','g','SI'],['Patilla','🍎 Frutas','u','SI'],
    ['Pera','🍎 Frutas','u','SI'],['Piña Golden','🍎 Frutas','u','SI'],
    ['Piña Lebrija','🍎 Frutas','u','SI'],['Pithaya','🍎 Frutas','u','SI'],
    ['Tomate de Árbol','🍎 Frutas','u','SI'],['Uchuva','🍎 Frutas','g','SI'],
    ['Uva Nacional','🍎 Frutas','g','SI'],['Uva Importada','🍎 Frutas','g','SI'],
    ['Pulpa Piña','🥤 Pulpas','u','SI'],['Pulpa Guayaba','🥤 Pulpas','u','SI'],
    ['Pulpa Guanábana','🥤 Pulpas','u','SI'],['Pulpa Mango','🥤 Pulpas','u','SI'],
    ['Pulpa Fresa','🥤 Pulpas','u','SI'],['Pulpa Maracuyá','🥤 Pulpas','u','SI'],
    ['Pulpa Mora','🥤 Pulpas','u','SI'],['Pulpa Lulo','🥤 Pulpas','u','SI'],
    ['Salpicón','🥤 Pulpas','u','SI']
  ];
  s.getRange(2,1,inicial.length,4).setValues(inicial);
  s.autoResizeColumns(1,4);
}
function getCatalogo(ss){
  const s=ss.getSheetByName(SH.CATALOGO);if(!s)return{};
  const rows=s.getDataRange().getValues().slice(1);
  const cat={};
  rows.filter(r=>String(r[3]).toUpperCase()==='SI'&&r[0]).forEach(r=>{
    const c=r[1]||'🛒 Otros';
    if(!cat[c])cat[c]=[];
    cat[c].push({n:String(r[0]),u:String(r[2]).toLowerCase()||'u'});
  });
  return cat;
}
function agregarProductoAlCatalogo(ss,data){
  const s=ss.getSheetByName(SH.CATALOGO);if(!s)return;
  s.appendRow([data.nombre,data.categoria||'🛒 Otros',data.tipo||'u','SI']);
}

// ── HELPERS ────────────────────────────────────
function encabezado(s,hdrs,color){
  s.appendRow(hdrs);
  const r=s.getRange(1,1,1,hdrs.length);
  r.setBackground(color);r.setFontColor('#fff');r.setFontWeight('bold');
}
function fmt(d){return Utilities.formatDate(d,Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm');}
function fmtD(d){return Utilities.formatDate(d,Session.getScriptTimeZone(),'dd/MM/yyyy');}
function addMeses(d,m){const r=new Date(d);r.setMonth(r.getMonth()+m);return r;}
