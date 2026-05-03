// ═══════════════════════════════════════════════════════════
// SMART SHOPPER — Apps Script v3
// 6 pestañas: Pedidos · Clientes · Suscripciones · Ventas · Análisis · Catálogo
// Deploy > New deployment > Web app > Anyone (ejecutar como Tú)
// ═══════════════════════════════════════════════════════════

// ── NOMBRES DE HOJAS ───────────────────────────
const SH = {
  PEDIDOS:       'Pedidos',
  CLIENTES:      'Clientes',
  SUSCRIPCIONES: 'Suscripciones',
  VENTAS:        'Ventas',
  ANALISIS:      'Análisis de Productos',
  CATALOGO:      'Catálogo'
};

// ── COLUMNAS Pedidos (base 0) ──────────────────
const PC = {
  FECHA:0,CODIGO:1,NOMBRE:2,CELULAR:3,LOCALIDAD:4,BARRIO:5,
  DIRECCION:6,APTO:7,REFERENCIA:8,GPS:9,HORARIO:10,
  PRODUCTOS:11,NOTA:12,ITEMS:13,ESTADO:14,VALOR:15,YAPA:16,DOM_REST:17,COMISION:18
};

const COLORES = {
  'Lista de compra':'#ffffff','Por consignación':'#fff9c4',
  'Pagado':'#c8e6c9','Empacado':'#b3e5fc',
  'Entregado':'#e1bee7'
};

// ════════════════════════════════════════════════
//  doPost — recibe pedidos y actualizaciones
// ════════════════════════════════════════════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    inicializarHojas(ss);

    switch(data.accion) {
      case 'actualizar':    actualizarPedido(ss, data);   break;
      case 'suscripcion':   guardarSuscripcion(ss, data); break;
      case 'comision':      registrarComision(ss, data);  break;
      case 'agregarProducto': agregarProductoAlCatalogo(ss,data); break;
      default:
        escribirPedido(ss, data);
        actualizarCliente(ss, data);
        actualizarAnalisis(ss, data);
    }
    return jsonOut({ ok: true });
  } catch(err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// ════════════════════════════════════════════════
//  doGet — sirve datos al admin y a la app
// ════════════════════════════════════════════════
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  if (!p.accion) return txt('Smart Shopper v3 OK');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  inicializarHojas(ss);

  switch(p.accion) {
    case 'pedidos':       return jsonOut({ pedidos:      getPedidos(ss, p.estado) });
    case 'clientes':      return jsonOut({ clientes:     getClientes(ss) });
    case 'suscripciones': return jsonOut({ suscripciones: getSuscripciones(ss) });
    case 'ventas':        return jsonOut({ ventas:       getVentas(ss) });
    case 'analisis':      return jsonOut({ analisis:     getAnalisis(ss) });
    case 'catalogo':      return jsonOut({ catalogo:     getCatalogo(ss) });
    default:              return jsonOut({ error: 'Acción desconocida' });
  }
}

// ════════════════════════════════════════════════
//  INICIALIZAR — crea todas las hojas si no existen
// ════════════════════════════════════════════════
function inicializarHojas(ss) {
  crearPedidos(ss);
  crearClientes(ss);
  crearSuscripciones(ss);
  crearVentas(ss);
  crearAnalisis(ss);
  crearCatalogo(ss);
}

// ── PEDIDOS ────────────────────────────────────
function crearPedidos(ss) {
  if (ss.getSheetByName(SH.PEDIDOS)) return;
  const s = ss.insertSheet(SH.PEDIDOS);
  const h = ['Fecha','Código','Nombre','Celular','Localidad','Barrio','Dirección',
    'Apto','Referencia','GPS','Horario','Productos','Nota','Items',
    'Estado','Valor ($)','Yapa','Dom.Rest.','Comisión Plaza'];
  encabezado(s, h, '#1b4332');
  s.setFrozenRows(1); s.setColumnWidth(12, 320); s.setColumnWidth(7,200);
}

function escribirPedido(ss, d) {
  const s = ss.getSheetByName(SH.PEDIDOS);
  const fecha = fmt(new Date(d.timestamp||new Date()));
  s.appendRow([fecha,d.codigo||'',d.nombre,d.celular,
    d.localidad||'',d.barrio||'',d.direccion||'',d.apto||'',
    d.referencia||'',d.gps||'',d.horario||'',d.productos||'',
    d.nota||'',d.total_items||0,'Lista de compra','','',d.domicilios_restantes||0,'']);
  const lr=s.getLastRow();
  s.getRange(lr,1,1,19).setBackground(lr%2===0?'#f0faf2':'#ffffff');
}

function actualizarPedido(ss, d) {
  const s = ss.getSheetByName(SH.PEDIDOS); if(!s)return;
  const rows = s.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][PC.CODIGO])===String(d.codigo)){
      const f=i+1;
      if(d.estado!==undefined)  s.getRange(f,PC.ESTADO+1).setValue(d.estado);
      if(d.valor!==undefined)   s.getRange(f,PC.VALOR+1).setValue(d.valor);
      if(d.yapa!==undefined)    s.getRange(f,PC.YAPA+1).setValue(d.yapa);
      if(d.comision!==undefined)s.getRange(f,PC.COMISION+1).setValue(d.comision);
      const color=COLORES[d.estado]||'#ffffff';
      s.getRange(f,1,1,19).setBackground(color);
      // Si se marca Entregado → registrar en Ventas
      if(d.estado==='Entregado' && d.valor)
        registrarVentaDesde(ss,{...rows[i-1],...d});
      break;
    }
  }
  if(d.domicilios!==undefined && d.celular) actualizarDomCliente(ss,d.celular,d.domicilios);
}

function getPedidos(ss, estado) {
  const s = ss.getSheetByName(SH.PEDIDOS); if(!s) return [];
  const rows = s.getDataRange().getValues();
  return rows.slice(1).map((r,i)=>({
    fila:i+2, fecha:String(r[PC.FECHA]), codigo:r[PC.CODIGO],
    nombre:r[PC.NOMBRE], celular:r[PC.CELULAR], localidad:r[PC.LOCALIDAD],
    barrio:r[PC.BARRIO], direccion:r[PC.DIRECCION], apto:r[PC.APTO],
    referencia:r[PC.REFERENCIA], gps:r[PC.GPS], horario:r[PC.HORARIO],
    productos:r[PC.PRODUCTOS], nota:r[PC.NOTA], items:r[PC.ITEMS],
    estado:r[PC.ESTADO]||'Lista de compra', valor:r[PC.VALOR]||'',
    yapa:r[PC.YAPA]||'', domRestantes:r[PC.DOM_REST]||0, comision:r[PC.COMISION]||''
  })).filter(p=>!estado||p.estado===estado);
}

// ── CLIENTES ───────────────────────────────────
function crearClientes(ss) {
  if (ss.getSheetByName(SH.CLIENTES)) return;
  const s = ss.insertSheet(SH.CLIENTES);
  encabezado(s,['Celular','Nombre','Dirección','Primer Pedido','Último Pedido','Total Pedidos','Domicilios Disponibles'],'#2d6a4f');
  s.setFrozenRows(1);
}
function actualizarCliente(ss, d) {
  const s=ss.getSheetByName(SH.CLIENTES);if(!s)return;
  const cel=String(d.celular),rows=s.getDataRange().getValues();
  let f=-1;
  for(let i=1;i<rows.length;i++){if(String(rows[i][0])===cel){f=i+1;break;}}
  const ahora=fmt(new Date());
  if(f===-1){s.appendRow([cel,d.nombre,d.direccion||'',ahora,ahora,1,d.domicilios_restantes||0]);}
  else{
    s.getRange(f,2).setValue(d.nombre);
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
    celular:r[0],nombre:r[1],direccion:r[2],primerPedido:r[3],
    ultimoPedido:r[4],totalPedidos:r[5],domicilios:r[6]||0
  }));
}

// ── SUSCRIPCIONES ──────────────────────────────
function crearSuscripciones(ss) {
  if(ss.getSheetByName(SH.SUSCRIPCIONES))return;
  const s=ss.insertSheet(SH.SUSCRIPCIONES);
  encabezado(s,['Celular','Nombre','Plan','Valor Pagado','Fecha Pago',
    'Domicilios Totales','Domicilios Usados','Domicilios Restantes',
    'Estado','Vence'],'#1565c0');
  s.setFrozenRows(1);
  // Fórmula automática de restantes en col H
  s.getRange('H2').setFormula('=IF(F2="","",F2-G2)');
}
function guardarSuscripcion(ss,d){
  const s=ss.getSheetByName(SH.SUSCRIPCIONES);
  const cel=String(d.celular),rows=s.getDataRange().getValues();
  let f=-1;
  for(let i=1;i<rows.length;i++){if(String(rows[i][0])===cel){f=i+1;break;}}
  const fecha=fmt(new Date(d.fechaPago||new Date()));
  const vence=d.meses>0?fmtFecha(addMeses(new Date(),d.meses)):'N/A';
  if(f===-1){
    s.appendRow([cel,d.nombre,d.plan,d.valor,fecha,d.envios,0,d.envios,'Activo',vence]);
  } else {
    // Sumar envíos al plan existente
    const domAct=Number(rows[f-1][7])||0;
    const domTot=Number(rows[f-1][5])||0;
    s.getRange(f,4).setValue(d.valor);
    s.getRange(f,5).setValue(fecha);
    s.getRange(f,6).setValue(domTot+d.envios);
    s.getRange(f,8).setValue(domAct+d.envios);
    s.getRange(f,9).setValue('Activo');
    s.getRange(f,10).setValue(vence);
  }
  actualizarDomCliente(ss,cel,(Number(rows[f>-1?f-1:rows.length-1]?.[7])||0)+d.envios);
}
function getSuscripciones(ss){
  const s=ss.getSheetByName(SH.SUSCRIPCIONES);if(!s)return[];
  return s.getDataRange().getValues().slice(1).map(r=>({
    celular:r[0],nombre:r[1],plan:r[2],valorPagado:r[3],
    fechaPago:r[4],domTotales:r[5],domUsados:r[6],
    domRestantes:r[7],estado:r[8],vence:r[9]
  }));
}

// ── VENTAS Y COMISIONES ────────────────────────
function crearVentas(ss){
  if(ss.getSheetByName(SH.VENTAS))return;
  const s=ss.insertSheet(SH.VENTAS);
  encabezado(s,['Fecha','Código Pedido','Cliente','Celular','Localidad',
    'Valor Pedido ($)','Comisión Plaza ($)','% Plaza',
    'Ingreso Neto Smart Shopper ($)','Mes'],'#880e4f');
  s.setFrozenRows(1);
  // Fila de totales
  s.getRange('A1').setNote('Registra automáticamente cuando un pedido se marca como Entregado');
}
function registrarVentaDesde(ss, row) {
  const s=ss.getSheetByName(SH.VENTAS);if(!s)return;
  const valor=Number(row.valor||row[PC.VALOR])||0;
  const comision=Number(row.comision||row[PC.COMISION])||0;
  const mes=new Date().toLocaleDateString('es-CO',{month:'long',year:'numeric'});
  s.appendRow([
    fmt(new Date()),
    row.codigo||row[PC.CODIGO],
    row.nombre||row[PC.NOMBRE],
    row.celular||row[PC.CELULAR],
    row.localidad||row[PC.LOCALIDAD],
    valor, comision,
    valor>0?Math.round(comision/valor*100)+'%':'',
    valor-comision, mes
  ]);
  // Color alternado
  const lr=s.getLastRow();
  if(lr%2===0)s.getRange(lr,1,1,10).setBackground('#fce4ec');
}
function registrarComision(ss,d){
  // Actualiza comisión manualmente desde admin
  actualizarPedido(ss,{...d,comision:d.comision});
}
function getVentas(ss){
  const s=ss.getSheetByName(SH.VENTAS);if(!s)return{filas:[],totales:{}};
  const rows=s.getDataRange().getValues().slice(1);
  const filas=rows.map(r=>({
    fecha:r[0],codigo:r[1],cliente:r[2],celular:r[3],localidad:r[4],
    valor:r[5],comision:r[6],pctPlaza:r[7],neto:r[8],mes:r[9]
  }));
  // Totales
  const totales={
    totalVentas:filas.reduce((a,r)=>a+Number(r.valor||0),0),
    totalComisiones:filas.reduce((a,r)=>a+Number(r.comision||0),0),
    totalNeto:filas.reduce((a,r)=>a+Number(r.neto||0),0),
    numPedidos:filas.length
  };
  return{filas,totales};
}

// ── ANÁLISIS DE PRODUCTOS ──────────────────────
function crearAnalisis(ss){
  if(ss.getSheetByName(SH.ANALISIS))return;
  const s=ss.insertSheet(SH.ANALISIS);
  encabezado(s,['Producto','Categoría','Total Pedidos','Última Vez','Este Mes',
    'Mes Anterior','Tendencia','Unidad'],'#1a5c1a');
  s.setFrozenRows(1);
  s.getRange('A1').setNote('Se actualiza automáticamente con cada nuevo pedido');
}
function actualizarAnalisis(ss, data) {
  const s=ss.getSheetByName(SH.ANALISIS);if(!s)return;
  if(!data.productos)return;
  const prods=data.productos.split('|').map(p=>p.trim()).filter(Boolean);
  const rows=s.getDataRange().getValues();
  const idx={};// nombre → fila (base 1)
  for(let i=1;i<rows.length;i++)idx[rows[i][0]]=i+1;
  const mes=new Date().getMonth();
  prods.forEach(p=>{
    const nombre=p.split(':')[0].trim();
    if(!nombre)return;
    if(idx[nombre]){
      const f=idx[nombre];
      const tot=Number(s.getRange(f,3).getValue())||0;
      const esteMes=Number(s.getRange(f,5).getValue())||0;
      s.getRange(f,3).setValue(tot+1);
      s.getRange(f,4).setValue(fmt(new Date()));
      s.getRange(f,5).setValue(esteMes+1);
      // Tendencia simple
      const ant=Number(s.getRange(f,6).getValue())||0;
      s.getRange(f,7).setValue(esteMes+1>ant?'↑ Sube':'↓ Baja');
    } else {
      // Nuevo producto
      s.appendRow([nombre,'',1,fmt(new Date()),1,0,'Nuevo','']);
    }
  });
  // Ordenar por total desc
  if(s.getLastRow()>2)
    s.getRange(2,1,s.getLastRow()-1,8).sort({column:3,ascending:false});
}
function getAnalisis(ss){
  const s=ss.getSheetByName(SH.ANALISIS);if(!s)return[];
  return s.getDataRange().getValues().slice(1).map(r=>({
    producto:r[0],categoria:r[1],total:r[2],ultimaVez:r[3],
    esteMes:r[4],mesAnterior:r[5],tendencia:r[6],unidad:r[7]
  }));
}

// ── CATÁLOGO DINÁMICO ──────────────────────────
function crearCatalogo(ss){
  if(ss.getSheetByName(SH.CATALOGO))return;
  const s=ss.insertSheet(SH.CATALOGO);
  encabezado(s,['Nombre del Producto','Categoría','Tipo (u=unidad / g=gramos)','Activo (SI/NO)'],'#4a148c');
  s.setFrozenRows(1);
  s.getRange('A1').setNote('Agrega productos aquí. Los cambios se reflejan automáticamente en la app.');
  // Poblar catálogo inicial
  const inicial=[
    ['Acelga','🥦 Verduras','g','SI'],['Auyama','🥦 Verduras','g','SI'],
    ['Alcachofa','🥦 Verduras','u','SI'],['Ají','🥦 Verduras','u','SI'],
    ['Ajo Importado','🥦 Verduras','u','SI'],['Ajo Nacional','🥦 Verduras','u','SI'],
    ['Apio','🥦 Verduras','u','SI'],['Aromáticas','🥦 Verduras','u','SI'],
    ['Arracacha','🥦 Verduras','g','SI'],['Arveja Cáscara','🥦 Verduras','g','SI'],
    ['Arveja Desgranada','🥦 Verduras','g','SI'],['Brócoli','🥦 Verduras','u','SI'],
    ['Berenjena','🥦 Verduras','u','SI'],['Cabezona Blanca','🥦 Verduras','u','SI'],
    ['Cabezona Roja','🥦 Verduras','u','SI'],['Calabaza','🥦 Verduras','g','SI'],
    ['Camote','🥦 Verduras','g','SI'],['Cebolla Larga','🥦 Verduras','u','SI'],
    ['Cebolla Puerro','🥦 Verduras','u','SI'],['Cebollin','🥦 Verduras','u','SI'],
    ['Champiñón','🥦 Verduras','g','SI'],['Cilantro','🥦 Verduras','u','SI'],
    ['Coliflor','🥦 Verduras','u','SI'],['Espárragos','🥦 Verduras','g','SI'],
    ['Espinaca','🥦 Verduras','g','SI'],['Espinaca Baby','🥦 Verduras','g','SI'],
    ['Frijol','🥦 Verduras','g','SI'],['Guatila','🥦 Verduras','u','SI'],
    ['Habichuela','🥦 Verduras','g','SI'],['Jengibre','🥦 Verduras','g','SI'],
    ['Lechuga Cogollo','🥦 Verduras','u','SI'],['Lechuga Batavia','🥦 Verduras','u','SI'],
    ['Lechuga Morada','🥦 Verduras','u','SI'],['Lechuga Romana','🥦 Verduras','u','SI'],
    ['Lechuga Rúgula','🥦 Verduras','g','SI'],['Mazorca','🥦 Verduras','u','SI'],
    ['Mazorca Desgranada','🥦 Verduras','g','SI'],['Papa Criolla','🥦 Verduras','g','SI'],
    ['Papa Sabanera','🥦 Verduras','g','SI'],['Papa Pastusa','🥦 Verduras','g','SI'],
    ['Pepino Cohombro','🥦 Verduras','u','SI'],['Pepino Común','🥦 Verduras','u','SI'],
    ['Perejil','🥦 Verduras','u','SI'],['Pimentón','🥦 Verduras','u','SI'],
    ['Rábano','🥦 Verduras','u','SI'],['Remolacha','🥦 Verduras','g','SI'],
    ['Repollo Blanco','🥦 Verduras','u','SI'],['Repollo Morado','🥦 Verduras','u','SI'],
    ['Tomate Cherry','🥦 Verduras','g','SI'],['Tomate Chonto','🥦 Verduras','g','SI'],
    ['Tomate Milano','🥦 Verduras','g','SI'],['Yuca','🥦 Verduras','g','SI'],
    ['Zanahoria','🥦 Verduras','g','SI'],['Zanahoria Baby','🥦 Verduras','g','SI'],
    ['Zuchini','🥦 Verduras','u','SI'],
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
  if(inicial.length>0)s.getRange(2,1,inicial.length,4).setValues(inicial);
  s.autoResizeColumns(1,4);
}
function getCatalogo(ss){
  const s=ss.getSheetByName(SH.CATALOGO);if(!s)return{};
  const rows=s.getDataRange().getValues().slice(1);
  const cat={};
  rows.filter(r=>String(r[3]).toUpperCase()==='SI'&&r[0]).forEach(r=>{
    const categoria=r[1]||'🛒 Otros';
    if(!cat[categoria])cat[categoria]=[];
    cat[categoria].push({n:String(r[0]),u:String(r[2]).toLowerCase()||'u'});
  });
  return cat;
}

// ── HELPERS ────────────────────────────────────
function encabezado(s,hdrs,color){
  s.appendRow(hdrs);
  const r=s.getRange(1,1,1,hdrs.length);
  r.setBackground(color);r.setFontColor('#ffffff');r.setFontWeight('bold');
}
function fmt(d){return Utilities.formatDate(d,Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm');}
function fmtFecha(d){return Utilities.formatDate(d,Session.getScriptTimeZone(),'dd/MM/yyyy');}
function addMeses(d,m){const r=new Date(d);r.setMonth(r.getMonth()+m);return r;}
function jsonOut(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function txt(t){return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT);}

// ── Agregar producto al catálogo desde admin ───
function agregarProductoAlCatalogo(ss, data) {
  const s = ss.getSheetByName(SH.CATALOGO);
  if (!s) return;
  s.appendRow([data.nombre, data.categoria||'🛒 Otros', data.tipo||'u', 'SI']);
  s.autoResizeColumn(1);
}
