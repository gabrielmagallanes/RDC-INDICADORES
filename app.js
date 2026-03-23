// ═══════════════════════════════════════════════
// DATOS — Excel RESUMEN_AVANCE
// ═══════════════════════════════════════════════
var tableData = {
  // 1. AVANCE DE VENTAS
  // Av.Lineal = Avance/Presupuesto · Proy.Cierre = Proyección/Presupuesto (auto-calculados)
  avance: { cols:['Asesor','Presupuesto','Avance','Av. Lineal','Proyección','Proy. Cierre'], rows:[
    ['HORECA 1',   '496,275.88','271,884.72','54.8%','526,776.64','106.1%'],
    ['HORECA 3',   '354,346.88','161,567.22','45.6%','313,036.49', '88.3%'],
    ['HORECA 4',   '203,790.61','106,201.43','52.1%','205,765.27','101.0%'],
    ['COMERCIO 2', '658,390.39','416,265.13','63.2%','806,513.69','122.5%'],
    ['COMERCIO 5', '405,240.89','240,674.53','59.4%','466,306.90','115.1%'],
    ['PRESTADOR S.','—','—','—','—','—']
  ]},
  // 2. FRESH FOOD
  freshfood: { cols:['Asesor','Presupuesto','Avance','Av. Lineal','Proyección','Proy. Cierre'], rows:[
    ['HORECA 1',   '129,031.73','66,340.80','51.4%','128,535.30', '99.6%'],
    ['HORECA 3',    '51,857.28','27,550.95','53.1%', '53,379.97','102.9%'],
    ['HORECA 4',    '40,379.80','21,118.29','52.3%', '40,916.69','101.3%'],
    ['COMERCIO 2',   '7,000.00', '3,564.97','50.9%',  '6,907.13', '98.7%'],
    ['COMERCIO 5',   '7,000.00', '1,172.10','16.7%',  '2,270.94', '32.4%'],
    ['PRESTADOR S.','—','—','—','—','—']
  ]},
  // 3. MMPP
  mmpp: { cols:['Asesor','Presupuesto','Avance','Av. Lineal','Proyección','Proy. Cierre'], rows:[
    ['HORECA 1',   '173,696.56', '94,109.01','54.2%','182,336.21','105.0%'],
    ['HORECA 3',    '94,627.47', '48,145.66','50.9%', '93,282.22', '98.6%'],
    ['HORECA 4',    '54,554.39', '29,379.14','53.9%', '56,922.08','104.3%'],
    ['COMERCIO 2',  '20,411.63', '10,918.13','53.5%', '21,153.88','103.6%'],
    ['COMERCIO 5',  '12,394.71',  '8,213.60','66.3%', '15,913.85','128.4%'],
    ['PRESTADOR S.','—','—','—','—','—']
  ]},
  // 4. CLIENTES META
  clientes: { cols:['Asesor','Ctd Clientes','Cuota 45%','Llegaron Meta','Avance vs Obj','% Meta'], rows:[
    ['HORECA 1',    '134','60.3', '9','-51.3', '6.7%'],
    ['HORECA 3',    '139','62.6','15','-47.6','10.8%'],
    ['HORECA 4',    '131','58.95','11','-47.95','8.4%'],
    ['COMERCIO 2',  '244','109.8','36','-73.8','14.8%'],
    ['COMERCIO 5',  '220','99.0', '45','-54.0','20.5%'],
    ['PRESTADOR S.','0','0','0','0','0%']
  ]},
  // 5. ACTIVACIÓN
  activacion: { cols:['Asesor','Cli Activ 2024','Cli Activ 2025','Cli Vtas 0','Menor S/.500','% Activos'], rows:[
    ['HORECA 1',    '110','110','24', '49','82.1%'],
    ['HORECA 3',    '105','106','33', '67','76.3%'],
    ['HORECA 4',     '84', '92','39', '83','70.2%'],
    ['COMERCIO 2',  '181','195','49','109','79.9%'],
    ['COMERCIO 5',  '147','167','53','118','75.9%'],
    ['PRESTADOR S.','0','0','0','0','0%']
  ]},
  // 6. A.A. — AÑO ANTERIOR
  aa: { cols:['Asesor','VTA 2024','VTA 2025','Diferencia','Variación'], rows:[
    ['HORECA 1',    '227,070.79','271,884.72','+44,813.93','+19.7%'],
    ['HORECA 3',    '150,256.15','161,567.22','+11,311.07', '+7.5%'],
    ['HORECA 4',     '63,875.23','106,201.43','+42,326.20','+66.3%'],
    ['COMERCIO 2',  '267,276.79','416,265.13','+148,988.34','+55.7%'],
    ['COMERCIO 5',  '140,292.90','240,674.53','+100,381.63','+71.5%'],
    ['PRESTADOR S.','0','0','0','0%']
  ]},
  // 7. MARGEN
  margen: { cols:['Asesor','MG 2024','MG 2025','Diferencia','Part. Margen'], rows:[
    ['HORECA 1',    '19,051.70','25,677.77', '+6,626.06', '+9.4%'],
    ['HORECA 3',     '9,254.98','13,631.39', '+4,376.41', '+8.4%'],
    ['HORECA 4',     '3,421.06', '6,022.37', '+2,601.31', '+5.7%'],
    ['COMERCIO 2', '-10,407.59','-12,724.05','-2,316.46', '-3.1%'],
    ['COMERCIO 5',  '-4,285.27', '-9,444.81','-5,159.54', '-3.9%'],
    ['PRESTADOR S.','0','0','0','0%']
  ]},
  // 8. VISITAS DIARIAS
  visitas: { cols:['Asesor','PRO','OBJ','Variación','Estado'], rows:[
    ['HORECA 1',    '0.95','1.00', '-5.0%','❌ No cumplido (95.0%)'],
    ['HORECA 3',    '0.88','1.00','-12.0%','❌ No cumplido (88.0%)'],
    ['HORECA 4',    '0.91','1.00', '-9.0%','❌ No cumplido (91.0%)'],
    ['COMERCIO 2',  '0.96','1.00', '-4.0%','❌ No cumplido (96.0%)'],
    ['COMERCIO 5',  '0.94','1.00', '-6.0%','❌ No cumplido (94.0%)'],
    ['PRESTADOR S.','0','0','—','—']
  ]},
  // 9. CONVERSIÓN
  conversion: { cols:['Asesor','PRO','OBJ','Variación','Estado'], rows:[
    ['HORECA 1',    '0.49','0.60','-18.3%','❌ No cumplido (81.7%)'],
    ['HORECA 3',    '0.35','0.60','-41.7%','❌ No cumplido (58.3%)'],
    ['HORECA 4',    '0.27','0.60','-55.0%','❌ No cumplido (45.0%)'],
    ['COMERCIO 2',  '0.27','0.50','-46.0%','❌ No cumplido (54.0%)'],
    ['COMERCIO 5',  '0.26','0.50','-48.0%','❌ No cumplido (52.0%)'],
    ['PRESTADOR S.','—','—','—','—']
  ]},
  // 10. PRODUCTIVIDAD — Estado se calcula automáticamente
  productividad: { cols:['Mes','PRO 2026','PRO 2025','OBJ','Estado'], rows:[
    ['ENERO',    '0.86','0.70','—',''],
    ['FEBRERO',  '0.86','0.70','—',''],
    ['MARZO',    '—',  '0.70','—',''],
    ['ABRIL',    '—',  '0.70','—',''],
    ['MAYO',     '—',  '0.70','—',''],
    ['JUNIO',    '—',  '0.70','—',''],
    ['JULIO',    '—',  '0.70','—',''],
    ['AGOSTO',   '—',  '0.70','—',''],
    ['SETIEMBRE','—',  '0.70','—',''],
    ['OCTUBRE',  '—',  '0.70','—',''],
    ['NOVIEMBRE','—',  '0.70','—',''],
    ['DICIEMBRE','—',  '0.70','—','']
  ]},
  // 11. CLIMA MANDELA
  clima: { cols:['Bimestre','PRO 2026','OBJ','Estado'], rows:[
    ['1 BIM','0.95','0.97','⚠ Bajo obj. (97.9%)'],
    ['2 BIM','—',  '0.97','Pendiente']
  ]},
  // 12. CDO PERFORMAS: 100 PUNTOS
  cdo: { cols:['Asesor','35 Ptos Cuota','10 Ptos Cli Meta','25 Ptos Venta OB','20 Ptos Venta FF','10 Ptos Cartera CA','Puntaje Total CDO'], rows:[
    ['HORECA 1',   '37.2','1.5','26.2','19.9','10','94.8'],
    ['HORECA 3',   '30.9','2.4','24.6','20.6','10','88.5'],
    ['HORECA 4',   '35.3','1.9','26.1','20.3','10','93.6'],
    ['COMERCIO 2', '55.1','8.2','20.7', '0', '12','96.0'],
    ['COMERCIO 5', '51.8','11.4','25.7','0',  '12','100.8'],
    ['PRESTADOR S.','0','0','0','0','0','0.0']
  ]}
};
var customTables = {}; // { key: { name, cols:[], rows:[], color } }
var ACCENT = ['#2e74e8','#e8622a','#17a55e','#9b4fd4','#c89a00','#0fa8cc'];
var chartInstances = {};
var selectedHex = '#2e74e8';
var activeView = 'tables';
var STORAGE_KEY = 'indicadoresRDC_v4';

// ═══════════════════════════════════════════════
// RELOJ
// ═══════════════════════════════════════════════
function updateClock(){
  var d=new Date();
  var dias=['Dom','Lun','Mar','Mie','Jue','Vie','Sab'];
  var meses=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Set','Oct','Nov','Dic'];
  var h=d.getHours().toString().padStart(2,'0'), m=d.getMinutes().toString().padStart(2,'0');
  var el=document.getElementById('sb-clock');
  if(el) el.textContent=dias[d.getDay()]+' '+d.getDate()+' '+meses[d.getMonth()]+' '+d.getFullYear()+'  '+h+':'+m;
}
updateClock();
setInterval(updateClock,30000);

// ═══════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════
function avg(key,col){ var td=getRows(key); var s=0,n=0; td.forEach(function(r){var v=parseFloat(r[col]);if(!isNaN(v)){s+=v;n++;}}); return n?s/n:0; }
function avgCol(rows,col){ var s=0,n=0; rows.forEach(function(r){var v=parseFloat(r[col]);if(!isNaN(v)){s+=v;n++;}}); return n?s/n:0; }
function getRows(key){ return tableData[key]?tableData[key].rows:(customTables[key]?customTables[key].rows:[]); }
function getCols(key){ return tableData[key]?tableData[key].cols:(customTables[key]?customTables[key].cols:[]); }
function calcEstado(pro,obj){ var p=parseFloat(pro),o=parseFloat(obj); if(isNaN(p)||isNaN(o)||o===0)return'—'; var pct=(p/o)*100; return pct>=100?'✅ Cumplido ('+pct.toFixed(1)+'%)':'❌ No cumplido ('+pct.toFixed(1)+'%)'; }
function calcVar(pro,obj){ var p=parseFloat(pro),o=parseFloat(obj); if(isNaN(p)||isNaN(o)||o===0)return'—'; var v=((p-o)/o)*100; return (v>=0?'+':'')+v.toFixed(1)+'%'; }
function pill(txt,ok){ return '<span style="font-weight:600;color:'+(ok?'var(--c3)':'var(--cr)')+'">'+txt+'</span>'; }
function stateClass(txt){ if(txt.includes('✅'))return'estado-ok'; if(txt.includes('❌'))return'estado-fail'; if(txt.includes('⚠'))return'estado-warn'; return''; }

// ═══════════════════════════════════════════════
// SIDEBAR STATS
// ═══════════════════════════════════════════════
function updateSbStats(){
  var c=document.getElementById('sb-stats'); if(!c)return;
  var avR2=tableData.avance.rows.filter(function(r){return r[0]!=='PRESTADOR S.';});
  var avT2=avR2.reduce(function(s,r){return s+(parseFloat(r[2].replace(/,/g,''))||0);},0);
  var avP2=avR2.reduce(function(s,r){return s+(parseFloat(r[1].replace(/,/g,''))||0);},0);
  var avPct3=avP2>0?((avT2/avP2)*100).toFixed(0)+'%':'—';
  var cvF2=tableData.conversion.rows.filter(function(r){return r[0]!=='PRESTADOR S.';});
  var vsF2=tableData.visitas.rows.filter(function(r){return r[0]!=='PRESTADOR S.';});
  var cdoF2=tableData.cdo.rows.filter(function(r){return r[0]!=='PRESTADOR S.';});
  var cdoAvg=(cdoF2.reduce(function(s,r){return s+(parseFloat(r[6])||0);},0)/cdoF2.length).toFixed(1);
  var conv2=avgCol(cvF2,1); var convObj2=avgCol(cvF2,2);
  var convPct2=convObj2>0?((conv2/convObj2)*100).toFixed(0)+'%':'—';
  var items=[
    {lbl:'Avance',    val:avPct3,  color:'#2e74e8'},
    {lbl:'Conversión',val:convPct2,color:'#9b4fd4'},
    {lbl:'CDO Prom.', val:cdoAvg,  color:'#c89a00'},
    {lbl:'Product.',  val:tableData.productividad.rows[0][1]||'—',color:'#17a55e'}
  ];
  c.innerHTML=items.map(function(s){
    return '<div class="stat-card"><div class="stat-val" style="color:'+s.color+'">'+s.val+'</div><div class="stat-lbl">'+s.lbl+'</div></div>';
  }).join('');
}

// ═══════════════════════════════════════════════
// NAVEGACIÓN
// ═══════════════════════════════════════════════
function navigateTo(view){
  activeView=view;
  document.querySelectorAll('.view').forEach(function(v){v.classList.remove('active');});
  document.getElementById('view-'+view).classList.add('active');
  document.querySelectorAll('.nav-btn').forEach(function(b){b.classList.remove('active');});
  document.querySelectorAll('.nav-btn').forEach(function(b){
    if(b.getAttribute('onclick')&&b.getAttribute('onclick').indexOf("'"+view+"'")!==-1) b.classList.add('active');
  });
  if(view==='dashboard'){ renderDashboard(); }
  if(view==='summary')  { renderSummary();   }
  updateSbStats();
}

function onActualizar(){
  renderTables();
  renderDashboard();
  renderSummary();
  updateSbStats();
  showToast('⟳ Vistas actualizadas correctamente');
}

// ═══════════════════════════════════════════════
// TABLAS
// ═══════════════════════════════════════════════
function calcProdEstado(pro26, pro25){
  var p=parseFloat(pro26), b=parseFloat(pro25);
  if(isNaN(p)||p===0) return '—';
  if(isNaN(b)||b===0) return '—';
  var pct=(p/b)*100;
  if(pct>=100) return '✅ Cumplido ('+pct.toFixed(1)+'%)';
  return '❌ No cumplido ('+pct.toFixed(1)+'%)';
}

function renderTables(){
  var body=document.getElementById('tables-body'); if(!body)return;
  body.innerHTML='';
  // ── Auto-calcular todas las tablas ────────────────────────
  // Productividad: Estado vs PRO 2025
  tableData.productividad.rows.forEach(function(r){
    r[4] = calcProdEstado(r[1], r[2]);
  });
  // Avance / Fresh Food / MMPP: Av.Lineal y Proy.Cierre
  ['avance','freshfood','mmpp'].forEach(function(key){
    tableData[key].rows.forEach(function(r){
      if(r[0]!=='PRESTADOR S.') calcAvanceRow(r);
    });
  });
  // Visitas / Conversión: Variación y Estado
  ['visitas','conversion'].forEach(function(key){
    tableData[key].rows.forEach(function(r){
      if(r[0]==='PRESTADOR S.'||r[0]==='PRESTADOR') return;
      var p=parseFloat(r[1]), o=parseFloat(r[2]);
      if(!isNaN(p)&&!isNaN(o)){
        r[3]=calcVar(r[1],r[2]);
        r[4]=calcEstado(r[1],r[2]);
      }
    });
  });
  // Clima Mandela: Estado
  tableData.clima.rows.forEach(function(r){
    if(r[1]&&r[1]!=='—') r[3]=calcEstado(r[1],r[2]);
    else r[3]='Pendiente';
  });
  // A.A.: Diferencia y Variación
  tableData.aa.rows.forEach(function(r){
    if(r[0]==='PRESTADOR S.') return;
    var v24=parseFloat(r[1].replace(/,/g,'')), v25=parseFloat(r[2].replace(/,/g,''));
    if(!isNaN(v24)&&!isNaN(v25)){
      var diff=v25-v24;
      r[3]=(diff>=0?'+':'')+diff.toLocaleString('es-PE',{maximumFractionDigits:2});
      r[4]=(diff>=0?'+':'')+((diff/v24)*100).toFixed(1)+'%';
    }
  });
  // Margen: Diferencia y Part.Margen
  tableData.margen.rows.forEach(function(r){
    if(r[0]==='PRESTADOR S.') return;
    var mg24=parseFloat(r[1].replace(/,/g,'')), mg25=parseFloat(r[2].replace(/,/g,''));
    if(!isNaN(mg24)&&!isNaN(mg25)){
      var diff=mg25-mg24;
      r[3]=(diff>=0?'+':'')+diff.toLocaleString('es-PE',{maximumFractionDigits:2});
      // Part.Margen = diff / abs(mg24)
      if(mg24!==0) r[4]=(diff>=0?'+':'')+((diff/Math.abs(mg24))*100).toFixed(1)+'%';
    }
  });
  // Clientes Meta: % Meta = Llegaron Meta / Cuota 45%
  tableData.clientes.rows.forEach(function(r){
    if(r[0]==='PRESTADOR S.') return;
    var llegaron=parseFloat(r[3]), cuota=parseFloat(r[2]);
    if(!isNaN(llegaron)&&!isNaN(cuota)&&cuota>0){
      r[4]=(llegaron-cuota).toFixed(1);          // Avance vs Obj
      r[5]=((llegaron/cuota)*100).toFixed(1)+'%'; // % Meta
    }
  });
  // Activación: % Activos = Cli Activ 2025 / (Cli Activ 2025 + Cli Vtas 0 + Menor S/.500) aprox
  // En realidad % Activos = (Cli Activ 2025 - Cli Vtas 0) / Cli Activ 2025
  tableData.activacion.rows.forEach(function(r){
    if(r[0]==='PRESTADOR S.') return;
    var act25=parseFloat(r[2]), vtas0=parseFloat(r[3]);
    if(!isNaN(act25)&&!isNaN(vtas0)&&act25>0){
      r[5]=((( act25 - vtas0)/act25)*100).toFixed(1)+'%';
    }
  });
  // CDO: Puntaje Total = suma cols 1 a 5
  tableData.cdo.rows.forEach(function(r){
    if(r[0]==='PRESTADOR S.') return;
    var suma=0;
    for(var i=1;i<=5;i++) suma+=parseFloat(r[i])||0;
    r[6]=suma.toFixed(1);
  });

  var base=[
    {key:'avance',       label:'AVANCE',                    color:'#2e74e8'},
    {key:'freshfood',    label:'FRESH FOOD',                color:'#17a55e'},
    {key:'mmpp',         label:'MMPP',                      color:'#9b4fd4'},
    {key:'clientes',     label:'CLIENTES META',             color:'#0fa8cc'},
    {key:'activacion',   label:'ACTIVACIÓN',                color:'#e8622a'},
    {key:'aa',           label:'A.A. — AÑO ANTERIOR',      color:'#c89a00'},
    {key:'margen',       label:'MARGEN',                    color:'#2e74e8'},
    {key:'visitas',      label:'VISITAS DIARIAS',           color:'#17a55e'},
    {key:'conversion',   label:'CONVERSIÓN',                color:'#e8622a'},
    {key:'productividad',label:'PRODUCTIVIDAD',             color:'#9b4fd4'},
    {key:'clima',        label:'CLIMA MANDELA',             color:'#0fa8cc'},
    {key:'cdo',          label:'CDO PERFORMAS: 100 PUNTOS', color:'#c89a00'}
  ];
  base.forEach(function(b){ body.appendChild(buildTableCard(b.key,b.label,b.color,tableData[b.key].cols,tableData[b.key].rows)); });
  Object.keys(customTables).forEach(function(k){
    var t=customTables[k];
    body.appendChild(buildTableCard(k,t.name,t.color,t.cols,t.rows));
  });
}

// Auto-calc column indices per table key (cannot be edited directly)
var AUTO_COLS = {
  avance:       {3:true, 5:true},   // Av.Lineal, Proy.Cierre
  freshfood:    {3:true, 5:true},
  mmpp:         {3:true, 5:true},
  clientes:     {4:true, 5:true},   // Avance vs Obj, % Meta
  activacion:   {5:true},           // % Activos
  aa:           {3:true, 4:true},   // Diferencia, Variación
  margen:       {3:true, 4:true},   // Diferencia, Part.Margen
  visitas:      {3:true, 4:true},   // Variación, Estado
  conversion:   {3:true, 4:true},
  productividad:{4:true},           // Estado
  clima:        {3:true},           // Estado
  cdo:          {6:true}            // Puntaje Total
};

function buildTableCard(key,label,color,cols,rows){
  var card=document.createElement('div'); card.className='table-card';
  var html='<div class="tc-hdr"><span class="tc-dot" style="background:'+color+'"></span><span class="tc-title">'+label+'</span></div>';
  var autoCols = AUTO_COLS[key] || {};
  var isCDO    = (key==='cdo');
  var isAvance = isAvanceTable(key);

  // Header — columnas auto-calculadas en azul con icono 🔄
  html+='<table><thead><tr>'+cols.map(function(c,ci){
    var isAuto=!!autoCols[ci];
    return '<th style="'+(isAuto?'color:#2e74e8':'')+'">'+c
      +(isAuto?' <span style="font-size:9px;opacity:.6">🔄</span>':'')+'</th>';
  }).join('')+'</tr></thead><tbody>';

  rows.forEach(function(r){
    var puntaje  = isCDO ? parseFloat(r[r.length-1]) : null;
    var rowBg    = (isCDO&&puntaje>=100)?'background:rgba(23,165,94,.08)':'';
    html+='<tr style="'+rowBg+'">'+r.map(function(cell,ci){
      var cls='', sty='';
      var isAuto = !!autoCols[ci];
      var isPS   = (r[0]==='PRESTADOR S.'||r[0]==='PRESTADOR');
      var pctVal = parseFloat(cell);

      if(isPS){ return '<td>'+cell+'</td>'; }

      if(isAuto){
        // Fondo azul claro — celda solo lectura visualmente
        sty='background:rgba(46,116,232,.05);font-weight:600';
        var sc=stateClass(cell);
        if(sc){ cls=sc; sty='background:rgba(46,116,232,.04);font-weight:600'; }
        // Colorear según tabla y columna
        else if((isAvance)&&(ci===3||ci===5)){
          cls=pctVal>=100?'pro-ok':pctVal>=50?'estado-warn':'pro-fail';
        } else if((key==='visitas'||key==='conversion')&&ci===3){
          cls=pctVal>=0?'estado-warn':'pro-fail'; // Variación
        } else if((key==='aa')&&ci===4){
          cls=pctVal>=0?'pro-ok':'pro-fail';
        } else if(key==='margen'&&ci===4){
          cls=pctVal>=0?'pro-ok':'pro-fail';
        } else if(key==='clientes'&&ci===5){
          cls=pctVal>=100?'pro-ok':pctVal>=50?'estado-warn':'pro-fail';
        } else if(key==='activacion'&&ci===5){
          cls=pctVal>=80?'pro-ok':pctVal>=70?'estado-warn':'pro-fail';
        } else if(key==='cdo'&&ci===6){
          cls=pctVal>=100?'pro-ok':pctVal>=90?'estado-warn':'pro-fail';
        } else if(key==='margen'&&ci===3){
          cls=pctVal>=0?'pro-ok':'pro-fail';
        } else if(key==='aa'&&ci===3){
          cls=pctVal>=0?'pro-ok':'pro-fail';
        } else {
          cls=(!isNaN(pctVal)&&pctVal<0)?'pro-fail':'';
        }
      } else {
        // Columnas editables: colorear según tabla
        if((key==='visitas'||key==='conversion')&&ci===1){
          var p=parseFloat(cell), o=parseFloat(r[2]||'0');
          if(!isNaN(p)&&!isNaN(o)&&o>0) cls=p>=o?'pro-ok':'pro-fail';
        } else if(key==='aa'&&ci===2){
          var v24=parseFloat(r[1].replace(/,/g,'')), v25=parseFloat(cell.replace(/,/g,''));
          if(!isNaN(v24)&&!isNaN(v25)) cls=v25>=v24?'pro-ok':'pro-fail';
        } else if(key==='margen'&&ci===2){
          var mg=parseFloat(cell.replace(/,/g,''));
          if(!isNaN(mg)) cls=mg>=0?'pro-ok':'pro-fail';
        } else {
          var sc2=stateClass(cell); if(sc2) cls=sc2;
        }
      }
      return '<td class="'+cls+'" style="'+sty+'">'+cell+'</td>';
    }).join('')+'</tr>';
  });
  html+='</tbody></table>';
  card.innerHTML=html; return card;
}

// ═══════════════════════════════════════════════
// DASHBOARD — KPIs
// ═══════════════════════════════════════════════
function renderDashboard(){
  renderKPIs();
  renderCharts();
}

function renderKPIs(){
  var g=document.getElementById('kpi-grid'); if(!g)return;
  // KPIs calculados desde datos reales del Excel
  var avRow=tableData.avance.rows; var avTotal=avRow.reduce(function(s,r){return s+parseFloat(r[2].replace(/,/g,''))||0;},0);
  var avPres=avRow.reduce(function(s,r){return s+parseFloat(r[1].replace(/,/g,''))||0;},0);
  var avPct=avPres>0?((avTotal/avPres)*100).toFixed(1)+'%':'—';
  var convRows=tableData.conversion.rows.filter(function(r){return r[0]!=='PRESTADOR S.';});
  var visRows=tableData.visitas.rows.filter(function(r){return r[0]!=='PRESTADOR S.';});
  var climRow=tableData.clima.rows[0];
  var items=[
    {lbl:'Avance Ventas',val:avPct,sub:'Presupuesto ejecutado',up:parseFloat(avPct)>=100,color:'#2e74e8'},
    {lbl:'Clima Mandela',val:climRow[1]||'—',sub:'OBJ: '+climRow[2],up:parseFloat(climRow[1])>=(parseFloat(climRow[2])||0.97),color:'#e8622a'},
    {lbl:'Productividad',val:tableData.productividad.rows[0][1]||'—',sub:'vs '+tableData.productividad.rows[0][2]+' (2025)',up:parseFloat(tableData.productividad.rows[0][1])>=parseFloat(tableData.productividad.rows[0][2]),color:'#17a55e'},
    {lbl:'Conversión prom',val:avgCol(convRows,1).toFixed(2),sub:'OBJ: '+avgCol(convRows,2).toFixed(2),up:avgCol(convRows,1)>=avgCol(convRows,2),color:'#9b4fd4'},
    {lbl:'Visitas prom',val:avgCol(visRows,1).toFixed(2),sub:'OBJ: '+avgCol(visRows,2).toFixed(2),up:avgCol(visRows,1)>=avgCol(visRows,2),color:'#c89a00'},
    {lbl:'Clientes Meta',val:tableData.clientes.rows.filter(function(r){return parseInt(r[3])>0;}).length+'/'+tableData.clientes.rows.filter(function(r){return r[0]!=='PRESTADOR S.';}).length,sub:'llegaron a meta',up:false,color:'#0fa8cc'}
  ];
  g.innerHTML=items.map(function(k){
    return '<div class="kpi-card" style="border-top-color:'+k.color+'">'
      +'<div class="kpi-lbl">'+k.lbl+'</div>'
      +'<div class="kpi-val" style="color:'+k.color+'">'+k.val+'</div>'
      +'<div class="kpi-sub">'+k.sub+'</div>'
      +'<div class="kpi-trend '+(k.up?'up':'down')+'">'+(k.up?'↑ Objetivo alcanzado':'↓ Bajo objetivo')+'</div>'
      +'</div>';
  }).join('');
}

// ═══════════════════════════════════════════════
// DASHBOARD — GRÁFICOS
// ═══════════════════════════════════════════════
function destroyChart(id){ if(chartInstances[id]){chartInstances[id].destroy();delete chartInstances[id];} }

function makeBarData(rows,proCol,objCol){
  var labels=[],pros=[],objs=[];
  rows.forEach(function(r){
    var p=parseFloat(r[proCol]),o=parseFloat(r[objCol]);
    if(!isNaN(p)){labels.push(r[0]);pros.push(p);objs.push(isNaN(o)?null:o);}
  });
  return{labels:labels,pros:pros,objs:objs};
}

function buildChartCard(id,title,sub){
  var card=document.createElement('div'); card.className='chart-card';
  card.innerHTML='<div class="chart-title">'+title+'</div>'
    +'<div class="chart-sub">'+sub+'</div>'
    +'<div class="chart-area"><canvas id="'+id+'"></canvas></div>';
  return card;
}

var tipCfg={backgroundColor:'#fff',borderColor:'#e2e6f0',borderWidth:1,titleColor:'#1a1f2e',bodyColor:'#8490a8',padding:10,cornerRadius:8};
var scaleCfg={x:{ticks:{color:'#8490a8',font:{size:11}},grid:{color:'rgba(226,230,240,.6)'},border:{dash:[4,4]}},y:{ticks:{color:'#8490a8',font:{size:11},callback:function(v){return v.toFixed(2);}},grid:{color:'rgba(226,230,240,.6)'},border:{dash:[4,4]}}};
var animCfg={duration:900,easing:'easeInOutQuart',delay:function(ctx){return ctx.dataIndex*60;}};

function renderBarChart(canvasId,rows,proCol,objCol,color){
  destroyChart(canvasId);
  var d=makeBarData(rows,proCol,objCol);
  var ctx=document.getElementById(canvasId); if(!ctx)return;
  chartInstances[canvasId]=new Chart(ctx,{type:'bar',data:{labels:d.labels,datasets:[
    {label:'PRO',data:d.pros,backgroundColor:d.pros.map(function(v,i){return(d.objs[i]!==null&&v>=d.objs[i])?'rgba(23,165,94,.75)':'rgba(204,51,51,.7)';}),borderColor:d.pros.map(function(v,i){return(d.objs[i]!==null&&v>=d.objs[i])?'#17a55e':'#cc3333';}),borderWidth:2,borderRadius:5,maxBarThickness:28},
    {label:'OBJ',data:d.objs,backgroundColor:'rgba(180,180,180,.18)',borderColor:'#ccc',borderWidth:1.5,borderRadius:5,maxBarThickness:28}
  ]},options:{responsive:true,maintainAspectRatio:false,animation:animCfg,plugins:{legend:{display:false},tooltip:tipCfg},scales:scaleCfg}});
}

// ── Chart card helper ───────────────────────────────────────────────────────
function mkCard(id, title, h){
  var d=document.createElement('div');
  d.className='mk-chart-card';
  d.innerHTML='<div class="mk-chart-title">'+title+'</div>'
    +'<div style="position:relative;height:'+(h||230)+'px"><canvas id="'+id+'"></canvas></div>';
  return d;
}

// ── Gráfico barras animado con gradiente ─────────────────────────────────────
function makroChart(id, labels, vals, objs){
  destroyChart(id);
  var ctx=document.getElementById(id); if(!ctx)return;
  var animDel={duration:800,easing:'easeInOutQuart',delay:function(c){return c.dataIndex*60;}};
  var tipCfg2={backgroundColor:'rgba(26,58,107,.93)',titleColor:'#fff',bodyColor:'#cce',
    padding:10,cornerRadius:8,
    callbacks:{label:function(c){return ' '+c.dataset.label+': '+c.raw+'%';}}};
  var datasets=[{
    type:'bar',label:'2026',data:vals,
    backgroundColor:function(c){
      var cv=c.chart.ctx; if(!cv)return '#1a3a6b';
      var grad=cv.createLinearGradient(0,0,0,220);
      grad.addColorStop(0,'#2e74e8'); grad.addColorStop(1,'#1a3a6b');
      return grad;
    },
    borderColor:'#1a3a6b',borderWidth:0,borderRadius:5,maxBarThickness:36,
    borderSkipped:false
  }];
  if(objs&&objs.length){
    datasets.push({
      type:'line',label:'OBJ',data:objs,
      borderColor:'#2ecc71',backgroundColor:'transparent',
      borderWidth:2.5,pointBackgroundColor:'#2ecc71',pointBorderColor:'#fff',
      pointBorderWidth:2,pointRadius:5,pointHoverRadius:8,
      tension:0,fill:false
    });
  }
  chartInstances[id]=new Chart(ctx,{
    data:{labels:labels,datasets:datasets},
    options:{responsive:true,maintainAspectRatio:false,animation:animDel,
      layout:{padding:{top:30,right:10}},
      plugins:{
        legend:{display:true,position:'bottom',
          labels:{color:'#444',font:{size:11},boxWidth:14,padding:14,usePointStyle:true}},
        tooltip:tipCfg2
      },
      scales:{
        x:{ticks:{color:'#555',font:{size:11},maxRotation:0},
           grid:{display:false},border:{display:false}},
        y:{display:true,position:'left',min:0,max:130,
           ticks:{color:'#a0aabb',font:{size:10},
             stepSize:20,callback:function(v){return v+'%';}},
           grid:{color:'rgba(210,218,235,.65)',lineWidth:1},
           border:{display:false}}
      }
    },
    plugins:[{
      id:'topLabels',
      afterDatasetsDraw:function(chart){
        var ctx2=chart.ctx;
        ctx2.save();
        chart.data.datasets.forEach(function(ds,di){
          var meta=chart.getDatasetMeta(di);
          meta.data.forEach(function(el,i){
            var val=ds.data[i];
            if(val===null||val===undefined) return;
            ctx2.font='bold 10px Inter,sans-serif';
            ctx2.fillStyle=di===0?'#1a3a6b':'#27ae60';
            ctx2.textAlign='center';
            var y=el.y-(di===0?6:18);
            ctx2.fillText(val+'%',el.x,y);
          });
        });
        ctx2.restore();
      }
    }]
  });
}

// ── Gráfico línea doble animado ──────────────────────────────────────────────
function makroLineChart(id, labels, data1, data2, lbl1, lbl2){
  destroyChart(id);
  var ctx=document.getElementById(id); if(!ctx)return;
  chartInstances[id]=new Chart(ctx,{type:'line',data:{labels:labels,datasets:[
    {label:lbl1,data:data1,borderColor:'#2e74e8',
     backgroundColor:function(c){
       var cv=c.chart.ctx,ch=c.chart.chartArea;
       if(!cv||!ch)return 'rgba(46,116,232,.1)';
       var g=cv.createLinearGradient(0,ch.top,0,ch.bottom);
       g.addColorStop(0,'rgba(46,116,232,.25)'); g.addColorStop(1,'rgba(46,116,232,.02)');
       return g;
     },
     tension:.35,fill:true,pointBackgroundColor:'#2e74e8',pointBorderColor:'#fff',
     pointBorderWidth:2,pointRadius:5,pointHoverRadius:8,borderWidth:2.5,spanGaps:false},
    {label:lbl2,data:data2,borderColor:'#f39c12',backgroundColor:'transparent',
     tension:.35,borderDash:[7,4],pointBackgroundColor:'#f39c12',pointBorderColor:'#fff',
     pointBorderWidth:2,pointRadius:4,pointHoverRadius:7,borderWidth:2,spanGaps:true}
  ]},options:{responsive:true,maintainAspectRatio:false,
    animation:{duration:900,easing:'easeInOutQuart'},layout:{padding:{top:10}},
    plugins:{
      legend:{display:true,position:'bottom',
        labels:{font:{size:11},boxWidth:14,padding:14,usePointStyle:true,color:'#444'}},
      tooltip:{backgroundColor:'rgba(26,58,107,.92)',titleColor:'#fff',bodyColor:'#cce',
        padding:10,cornerRadius:8}},
    scales:{
      x:{ticks:{color:'#555',font:{size:11}},grid:{color:'rgba(210,218,235,.5)'},border:{display:false}},
      y:{ticks:{color:'#a0aabb',font:{size:10}},
         grid:{color:'rgba(210,218,235,.65)'},border:{display:false}}
    }
  }});
}

function renderCharts(){
  var r1=document.getElementById('chart-row-1');
  var r2=document.getElementById('chart-row-2');
  var cw=document.getElementById('chart-cumpl-wrap');
  var ca=document.getElementById('custom-charts-area');
  if(!r1||!r2||!cw||!ca)return;
  r1.innerHTML=''; r2.innerHTML=''; cw.innerHTML=''; ca.innerHTML='';

  // ── Misma distribución que la imagen Makro: 2 cols × filas ──────────────
  // Fila 1: AVANCE + PRODUCTIVIDAD
  r1.appendChild(mkCard('ch-av',  'AVANCE DE CARTERAS'));
  r1.appendChild(mkCard('ch-prod','PRODUCTIVIDAD'));

  // Fila 2: CONVERSIÓN + VENTAS (FRESH FOOD)
  r2.appendChild(mkCard('ch-conv','CONVERSIÓN'));
  r2.appendChild(mkCard('ch-ff',  'FRESH FOOD'));

  // Fila 3: VISITAS + CLIMA MANDELA
  var r3=document.createElement('div'); r3.className='chart-row'; r3.style.marginBottom='14px';
  r3.appendChild(mkCard('ch-vis','VISITAS DIARIAS'));
  r3.appendChild(mkCard('ch-clim','PULSO — CLIMA MANDELA',180));
  ca.appendChild(r3);

  // Fila 4: MMPP + A.A.
  var r4=document.createElement('div'); r4.className='chart-row'; r4.style.marginBottom='14px';
  r4.appendChild(mkCard('ch-mmpp','MMPP'));
  r4.appendChild(mkCard('ch-aa',  'AÑO ANTERIOR — VARIACIÓN %'));
  ca.appendChild(r4);

  // Fila 5: ACTIVACIÓN + CLIENTES META
  var r5=document.createElement('div'); r5.className='chart-row'; r5.style.marginBottom='14px';
  r5.appendChild(mkCard('ch-act','ACTIVACIÓN DE CLIENTES'));
  r5.appendChild(mkCard('ch-cli','CLIENTES META'));
  ca.appendChild(r5);

  // Fila 6: MARGEN + CDO PERFORMAS
  var r6=document.createElement('div'); r6.className='chart-row'; r6.style.marginBottom='14px';
  r6.appendChild(mkCard('ch-mg', 'MARGEN'));
  r6.appendChild(mkCard('ch-cdo','CDO PERFORMAS — 100 PUNTOS'));
  ca.appendChild(r6);

  // Cumplimiento global
  var cumDiv=document.createElement('div'); cumDiv.className='mk-chart-card';
  cumDiv.innerHTML='<div class="mk-chart-title">CUMPLIMIENTO GLOBAL POR INDICADOR (%)</div>'
    +'<div style="position:relative;height:260px"><canvas id="ch-cumpl"></canvas></div>';
  cw.appendChild(cumDiv);

  var ASESORES=['HORECA 1','HORECA 3','HORECA 4','COMERCIO 2','COMERCIO 5'];
  function filt(key){ return tableData[key].rows.filter(function(r){return r[0]!=='PRESTADOR S.';}); }
  function pct(s){ return parseFloat(s.replace('%','').replace(',','')); }
  function num(s){ return parseFloat(s.replace(/,/g,'')); }

  setTimeout(function(){
    var avF  =filt('avance'), ffF=filt('freshfood'), mmF=filt('mmpp');
    var visF =filt('visitas'), convF=filt('conversion');
    var actF =filt('activacion'), cliF=filt('clientes');
    var aaF  =filt('aa'), mgF=filt('margen'), cdoF=filt('cdo');
    var prodR=tableData.productividad.rows;
    var climR=tableData.clima.rows;

    // ── AVANCE DE CARTERAS ──
    makroChart('ch-av', avF.map(function(r){return r[0];}),
      avF.map(function(r){return pct(r[3]);}),
      avF.map(function(){return 100;}));

    // ── PRODUCTIVIDAD ──
    var prodMeses=prodR.filter(function(r){return r[2]!=='—';}).map(function(r){return r[0];});
    var prod26=prodR.filter(function(r){return r[2]!=='—';}).map(function(r){return r[1]!=='—'?Math.round(pct(r[1])*100)/100:null;});
    var prod25=prodR.filter(function(r){return r[2]!=='—';}).map(function(r){return r[2]!=='—'?Math.round(num(r[2])*100)/100:null;});
    makroLineChart('ch-prod', prodMeses, prod26, prod25, 'PRO 2026', 'PRO 2025');

    // ── CONVERSIÓN ──
    makroChart('ch-conv', convF.map(function(r){return r[0];}),
      convF.map(function(r){return Math.round(pct(r[1])*100);}),
      convF.map(function(r){return Math.round(pct(r[2])*100);}));

    // ── FRESH FOOD ──
    makroChart('ch-ff', ffF.map(function(r){return r[0];}),
      ffF.map(function(r){return pct(r[3]);}),
      ffF.map(function(){return 100;}));

    // ── VISITAS DIARIAS ──
    makroChart('ch-vis', visF.map(function(r){return r[0];}),
      visF.map(function(r){return Math.round(pct(r[1])*100);}),
      visF.map(function(r){return Math.round(pct(r[2])*100);}));

    // ── CLIMA MANDELA ──
    makroChart('ch-clim', climR.map(function(r){return r[0];}),
      climR.map(function(r){return r[1]!=='—'?Math.round(pct(r[1])*100):null;}),
      climR.map(function(r){return r[2]!=='—'?Math.round(pct(r[2])*100):null;}));

    // ── MMPP ──
    makroChart('ch-mmpp', mmF.map(function(r){return r[0];}),
      mmF.map(function(r){return pct(r[3]);}),
      mmF.map(function(){return 100;}));

    // ── A.A. VARIACIÓN ──
    makroChart('ch-aa', aaF.map(function(r){return r[0];}),
      aaF.map(function(r){return pct(r[4]);}), null);

    // ── ACTIVACIÓN ──
    makroChart('ch-act', actF.map(function(r){return r[0];}),
      actF.map(function(r){return pct(r[5]);}),
      actF.map(function(){return 80;}));

    // ── CLIENTES META ──
    makroChart('ch-cli', cliF.map(function(r){return r[0];}),
      cliF.map(function(r){return pct(r[5]);}),
      cliF.map(function(){return 45;}));

    // ── MARGEN ──
    makroChart('ch-mg', mgF.map(function(r){return r[0];}),
      mgF.map(function(r){return pct(r[4]);}), null);

    // ── CDO PERFORMAS ──
    makroChart('ch-cdo', cdoF.map(function(r){return r[0];}),
      cdoF.map(function(r){return parseFloat(r[6]);}),
      cdoF.map(function(){return 100;}));

    renderCumplimiento();
  },60);
}

function renderProdChart(){
  destroyChart('ch-prod');
  var ctx=document.getElementById('ch-prod'); if(!ctx) return;
  var rows=tableData.productividad.rows;
  var labels=rows.map(function(r){return r[0];});
  var pro26=rows.map(function(r){var v=parseFloat(r[1]); return isNaN(v)?null:v;});
  var pro25=rows.map(function(r){var v=parseFloat(r[2]); return isNaN(v)?null:v;});
  chartInstances['ch-prod']=new Chart(ctx,{type:'line',data:{labels:labels,datasets:[
    {label:'PRO 2026',data:pro26,borderColor:'#17a55e',backgroundColor:'rgba(23,165,94,.1)',
     tension:.3,fill:true,pointBackgroundColor:'#17a55e',pointRadius:5,pointHoverRadius:7,
     borderWidth:2.5,spanGaps:false},
    {label:'PRO 2025',data:pro25,borderColor:'#8490a8',backgroundColor:'transparent',
     tension:.3,fill:false,pointBackgroundColor:'#8490a8',pointRadius:3,
     borderWidth:1.5,borderDash:[5,4],spanGaps:true}
  ]},options:{responsive:true,maintainAspectRatio:false,animation:animCfg,
    plugins:{legend:{display:true,labels:{color:'#1a1f2e',font:{size:11}}},tooltip:tipCfg},
    scales:scaleCfg}});
}

function renderClimaChart(){
  destroyChart('ch-clim');
  var ctx=document.getElementById('ch-clim'); if(!ctx) return;
  var rows=tableData.clima.rows;
  var labels=rows.map(function(r){return r[0];});
  var vals=rows.map(function(r){var v=parseFloat(r[1]); return isNaN(v)?0:v;});
  var objs=rows.map(function(r){var v=parseFloat(r[2]); return isNaN(v)?null:v;});
  chartInstances['ch-clim']=new Chart(ctx,{type:'bar',data:{labels:labels,datasets:[
    {label:'PRO 2026',data:vals,
     backgroundColor:vals.map(function(v,i){return(objs[i]&&v>=objs[i])?'rgba(23,165,94,.75)':'rgba(204,51,51,.7)';}),
     borderColor:vals.map(function(v,i){return(objs[i]&&v>=objs[i])?'#17a55e':'#cc3333';}),
     borderWidth:2,borderRadius:6,maxBarThickness:50},
    {label:'OBJ',data:objs,backgroundColor:'rgba(180,180,180,.2)',borderColor:'#ccc',
     borderWidth:1.5,borderRadius:6,maxBarThickness:50}
  ]},options:{responsive:true,maintainAspectRatio:false,animation:animCfg,
    plugins:{legend:{display:false},tooltip:tipCfg},scales:scaleCfg}});
}

function renderCumplimiento(){
  destroyChart('ch-cumpl');
  var ctx=document.getElementById('ch-cumpl'); if(!ctx)return;
  var visF=tableData.visitas.rows.filter(function(r){return r[0]!=='PRESTADOR S.';});
  var convF=tableData.conversion.rows.filter(function(r){return r[0]!=='PRESTADOR S.';});
  var avF=tableData.avance.rows.filter(function(r){return r[0]!=='PRESTADOR S.';});
  var avT=avF.reduce(function(s,r){return s+(parseFloat(r[2].replace(/,/g,''))||0);},0);
  var avP=avF.reduce(function(s,r){return s+(parseFloat(r[1].replace(/,/g,''))||0);},0);
  var avPct=avP>0?(avT/avP)*100:0;
  var prod26=parseFloat(tableData.productividad.rows[0][1])||0;
  var prod25=parseFloat(tableData.productividad.rows[0][2])||0.70;
  var clima1=parseFloat(tableData.clima.rows[0][1])||0;

  var labels=['Visitas','Conversión','Avance Vtas','Productividad','Clima Mandela','Activación'];
  var actF=tableData.activacion.rows.filter(function(r){return r[0]!=='PRESTADOR S.';});
  var actPct=actF.reduce(function(s,r){return s+parseFloat(r[5])||0;},0)/actF.length;

  var vals=[
    avgCol(visF,2)>0?(avgCol(visF,1)/avgCol(visF,2))*100:0,
    avgCol(convF,2)>0?(avgCol(convF,1)/avgCol(convF,2))*100:0,
    avPct,
    prod25>0?(prod26/prod25)*100:0,
    (clima1/0.97)*100,
    actPct
  ].map(function(v){return parseFloat((v||0).toFixed(1));});

  chartInstances['ch-cumpl']=new Chart(ctx,{type:'bar',data:{labels:labels,datasets:[{
    data:vals,
    backgroundColor:vals.map(function(v){return v>=100?'rgba(23,165,94,.78)':v>=75?'rgba(200,154,0,.75)':'rgba(204,51,51,.72)';}),
    borderColor:vals.map(function(v){return v>=100?'#17a55e':v>=75?'#c89a00':'#cc3333';}),
    borderWidth:2,borderRadius:8,maxBarThickness:70,label:'%'
  }]},options:{responsive:true,maintainAspectRatio:false,animation:animCfg,
    plugins:{legend:{display:false},tooltip:tipCfg},
    scales:{
      x:{ticks:{color:'#8490a8',font:{size:11}},grid:{color:'rgba(226,230,240,.6)'},border:{dash:[4,4]}},
      y:{min:0,max:130,ticks:{color:'#8490a8',font:{size:11},callback:function(v){return v+'%';}},
         grid:{color:'rgba(226,230,240,.6)'},border:{dash:[4,4]}}
    }
  }});
}

function renderCustomCharts(){
  var area=document.getElementById('custom-charts-area'); if(!area)return;
  area.innerHTML='';
  Object.keys(customTables).forEach(function(k){
    var t=customTables[k];
    var cid='ch-custom-'+k;
    var card=buildChartCard(cid,t.name,t.cols[1]+' vs '+t.cols[2]);
    card.style.marginBottom='16px';
    area.appendChild(card);
    setTimeout(function(){ renderBarChart(cid,t.rows,1,2,t.color); },80);
  });
}

// ═══════════════════════════════════════════════
// RESUMEN
// ═══════════════════════════════════════════════
function renderSummary(){
  var g=document.getElementById('res-cards'); if(!g)return;

  // ── Cálculos ──────────────────────────────────
  var noPS=function(key){ return tableData[key].rows.filter(function(r){return r[0]!=='PRESTADOR S.';}); };

  // Avance
  var avF=noPS('avance');
  var avTot=avF.reduce(function(s,r){return s+(parseFloat(r[2].replace(/,/g,''))||0);},0);
  var avPres=avF.reduce(function(s,r){return s+(parseFloat(r[1].replace(/,/g,''))||0);},0);
  var avPct=avPres>0?((avTot/avPres)*100).toFixed(1)+'%':'—';
  var avOk=parseFloat(avPct)>=50;

  // Fresh Food
  var ffF=noPS('freshfood');
  var ffTot=ffF.reduce(function(s,r){return s+(parseFloat(r[2].replace(/,/g,''))||0);},0);
  var ffPres=ffF.reduce(function(s,r){return s+(parseFloat(r[1].replace(/,/g,''))||0);},0);
  var ffPct=ffPres>0?((ffTot/ffPres)*100).toFixed(1)+'%':'—';
  var ffOk=parseFloat(ffPct)>=50;

  // MMPP
  var mmF=noPS('mmpp');
  var mmTot=mmF.reduce(function(s,r){return s+(parseFloat(r[2].replace(/,/g,''))||0);},0);
  var mmPres=mmF.reduce(function(s,r){return s+(parseFloat(r[1].replace(/,/g,''))||0);},0);
  var mmPct=mmPres>0?((mmTot/mmPres)*100).toFixed(1)+'%':'—';
  var mmOk=parseFloat(mmPct)>=50;

  // Clientes Meta
  var cliF=noPS('clientes');
  var cliTotal=cliF.reduce(function(s,r){return s+(parseInt(r[3])||0);},0);
  var cliCuota=cliF.reduce(function(s,r){return s+(parseInt(r[1])||0);},0);
  var cliOk=false;

  // Activación
  var actF=noPS('activacion');
  var actPct=(actF.reduce(function(s,r){return s+(parseFloat(r[5])||0);},0)/actF.length).toFixed(1)+'%';
  var actOk=parseFloat(actPct)>=80;

  // A.A.
  var aaF=noPS('aa');
  var aaCrecen=aaF.filter(function(r){return r[4].indexOf('+')===0;}).length;
  var aaOk=aaCrecen===aaF.length;

  // Margen
  var mgF=noPS('margen');
  var mgPos=mgF.filter(function(r){return parseFloat(r[2].replace(/,/g,''))>0;}).length;
  var mgOk=mgPos===mgF.length;

  // Visitas
  var visF=noPS('visitas');
  var visAvg=avgCol(visF,1).toFixed(2);
  var visObj=avgCol(visF,2).toFixed(2);
  var visOk=parseFloat(visAvg)>=parseFloat(visObj);

  // Conversión
  var convF=noPS('conversion');
  var convAvg=avgCol(convF,1).toFixed(2);
  var convObj=avgCol(convF,2).toFixed(2);
  var convOk=parseFloat(convAvg)>=parseFloat(convObj);

  // Productividad
  var prod26=parseFloat(tableData.productividad.rows[0][1])||0;
  var prod25=parseFloat(tableData.productividad.rows[0][2])||0.70;
  var prodPct=prod25>0?((prod26/prod25)*100).toFixed(1)+'%':'—';
  var prodOk=prod26>=prod25;

  // Clima
  var clima1=parseFloat(tableData.clima.rows[0][1])||0;
  var climaOk=clima1>=0.97;

  // CDO
  var cdoF=noPS('cdo');
  var cdoAvg=(cdoF.reduce(function(s,r){return s+(parseFloat(r[6])||0);},0)/cdoF.length).toFixed(1);
  var cdoMax=Math.max.apply(null,cdoF.map(function(r){return parseFloat(r[6])||0;}));
  var cdoOk=parseFloat(cdoAvg)>=90;

  // ── Cards en el orden solicitado ──────────────
  var cards=[
    {icon:'📊', title:'Avance de Ventas',
     l1:'Ejecutado: '+avPct+' del presupuesto',
     l2:'Total avance: S/. '+avTot.toLocaleString('es-PE',{maximumFractionDigits:0}),
     ok:avOk, color:'#2e74e8'},

    {icon:'🌿', title:'Fresh Food',
     l1:'Ejecutado: '+ffPct+' del presupuesto',
     l2:'COMERCIO 5: 16.7% — Requiere atención urgente',
     ok:ffOk, color:'#17a55e'},

    {icon:'🏗️', title:'MMPP',
     l1:'Ejecutado: '+mmPct+' del presupuesto',
     l2:'COMERCIO 5: 66.3% destaca · HORECA 3: 50.9%',
     ok:mmOk, color:'#9b4fd4'},

    {icon:'🎯', title:'Clientes Meta',
     l1:cliTotal+' clientes llegaron a meta de '+cliCuota+' cuota',
     l2:'COMERCIO 5: 45 · COMERCIO 2: 36 · HORECA 3: 15',
     ok:cliOk, color:'#0fa8cc'},

    {icon:'⚡', title:'Activación',
     l1:'% Activos promedio: '+actPct,
     l2:'HORECA 1: 82.1% · HORECA 3: 76.3% · HORECA 4: 70.2%',
     ok:actOk, color:'#e8622a'},

    {icon:'📅', title:'A.A. — Año Anterior',
     l1:aaCrecen+'/'+aaF.length+' asesores con crecimiento vs 2024',
     l2:'COMERCIO 2: +55.7%  ·  COMERCIO 5: +71.5%',
     ok:aaOk, color:'#c89a00'},

    {icon:'💰', title:'Margen',
     l1:mgPos+'/'+mgF.length+' asesores con margen positivo',
     l2:'COMERCIO 2: -3.1%  ·  COMERCIO 5: -3.9%',
     ok:mgOk, color:'#2e74e8'},

    {icon:'👣', title:'Visitas Diarias',
     l1:'Promedio: '+visAvg+' / OBJ: '+visObj,
     l2:'Todos ligeramente bajo el objetivo',
     ok:visOk, color:'#17a55e'},

    {icon:'⚠️', title:'Conversión',
     l1:'Promedio: '+convAvg+' / OBJ: '+convObj,
     l2:'Todos los asesores bajo el objetivo',
     ok:convOk, color:'#e8622a'},

    {icon:'📈', title:'Productividad',
     l1:'PRO 2026: '+(prod26||'—')+'  vs  PRO 2025: '+prod25,
     l2:'Cumplimiento: '+prodPct+' vs año anterior',
     ok:prodOk, color:'#9b4fd4'},

    {icon:'🌡️', title:'Clima Mandela',
     l1:'1er BIM: '+(tableData.clima.rows[0][1]||'—')+'  ·  OBJ: 0.97',
     l2:'2do BIM pendiente de datos',
     ok:climaOk, color:'#0fa8cc'},

    {icon:'🏆', title:'CDO Performas',
     l1:'Puntaje promedio: '+cdoAvg+' / 100 pts',
     l2:'Mejor: COMERCIO 5 con '+cdoMax.toFixed(1)+' pts · OBJ: 100',
     ok:cdoOk, color:'#c89a00'}
  ];

  g.innerHTML=cards.map(function(c){
    return '<div class="res-card" style="border-top-color:'+c.color+'">'
      +'<div class="res-icon">'+c.icon+'</div>'
      +'<div class="res-title">'+c.title+'</div>'
      +'<div class="res-l1">'+c.l1+'</div>'
      +'<div class="res-l2">'+c.l2+'</div>'
      +'<div class="res-status '+(c.ok?'res-ok':'res-fail')+'">'
      +(c.ok?'● En objetivo':'● Requiere atención')+'</div>'
      +'</div>';
  }).join('');

  // ── Semáforo con las 12 tablas ────────────────
  var sw=document.getElementById('semaforo-wrap'); if(!sw)return;
  var sem=[
    ['Avance Ventas',      avPct,               '≥50%',     parseFloat(avPct),                             avOk ?'🟢 EN CURSO':'🟡 REVISAR'],
    ['Fresh Food',         ffPct,               '≥50%',     parseFloat(ffPct),                             ffOk ?'🟢 EN CURSO':'🔴 BAJO'],
    ['MMPP',               mmPct,               '≥50%',     parseFloat(mmPct),                             mmOk ?'🟢 EN CURSO':'🔴 BAJO'],
    ['Clientes Meta',      cliTotal+' llegaron','Cuota 45%', null,                                          '🔴 BAJO OBJ.'],
    ['Activación',         actPct,              '≥80%',     parseFloat(actPct),                            actOk?'🟢 ÓPTIMO':'🟡 MONITOREAR'],
    ['A.A. Año Anterior',  aaCrecen+'/'+aaF.length+' crecen','Todos',aaCrecen/aaF.length*100,              aaOk ?'🟢 ÓPTIMO':'🟡 PARCIAL'],
    ['Margen',             mgPos+'/'+mgF.length+' positivos','Todos',mgPos/mgF.length*100,                 mgOk ?'🟢 ÓPTIMO':'🔴 CRÍTICO'],
    ['Visitas (prom.)',     visAvg,              visObj,     (parseFloat(visAvg)/parseFloat(visObj))*100,   '🟡 MONITOREAR'],
    ['Conversión (prom.)', convAvg,             convObj,    (parseFloat(convAvg)/parseFloat(convObj))*100, '🔴 CRÍTICO'],
    ['Productividad',      prod26||'—',         prod25+'', parseFloat(prodPct),                           prodOk ?'🟢 MEJORA':'🔴 BAJO'],
    ['Clima Mandela',      clima1||'—',         '0.97',     (clima1/0.97)*100,                             climaOk?'🟢 ÓPTIMO':'🟡 LEVE DESVÍO'],
    ['CDO Performas',      cdoAvg+' pts',       '100 pts',  parseFloat(cdoAvg),                            cdoOk ?'🟢 ÓPTIMO':'🟡 MONITOREAR']
  ];
  var semRows=sem.map(function(s){
    var sc=s[4].includes('🟢')?'estado-ok':s[4].includes('🔴')?'estado-fail':'estado-warn';
    var pct=s[3]!==null&&!isNaN(s[3])?s[3].toFixed(1)+'%':'—';
    return '<tr><td>'+s[0]+'</td><td style="font-weight:600">'+s[1]+'</td>'
      +'<td>'+s[2]+'</td><td>'+pct+'</td><td class="'+sc+'">'+s[4]+'</td></tr>';
  }).join('');
  sw.innerHTML='<div class="table-card"><div class="tc-hdr">'
    +'<span class="tc-dot" style="background:#2e74e8"></span>'
    +'<span class="tc-title">🚦 Semáforo de Indicadores — 12 Tablas</span></div>'
    +'<table><thead><tr><th>Indicador</th><th>Valor</th><th>Objetivo</th>'
    +'<th>Cumplimiento</th><th>Estado</th></tr></thead>'
    +'<tbody>'+semRows+'</tbody></table></div>';
}

// ═══════════════════════════════════════════════
// MODALES
// ═══════════════════════════════════════════════
function openModal(id){
  document.getElementById(id).classList.add('open');
  if(id==='modal-editcols') populateEditColsSel();
  if(id==='modal-addcol')   populateAddColSel();
  if(id==='modal-deltable') populateDeltableSel();
  if(id==='modal-delrow')   populateDelrowSel();
  if(id==='modal-addrow')   populateAddrowSel();
  if(id==='modal-edit')     populateEditSel();
}
function closeModal(id){
  document.getElementById(id).classList.remove('open');
  // Limpiar formularios al cerrar
  var cleanIds={'modal-edit':['edit-rows-container'],'modal-addrow':['addrow-form'],'modal-editcols':['editcols-form'],'modal-delrow':['delrow-form']};
  if(cleanIds[id]) cleanIds[id].forEach(function(cid){ var el=document.getElementById(cid); if(el) el.innerHTML=''; });
}
document.querySelectorAll('.modal-overlay').forEach(function(o){
  o.addEventListener('click',function(e){ if(e.target===o) o.classList.remove('open'); });
});

// Todas las tablas base + custom
var BASE_TABLES = [
  {key:'avance',       label:'AVANCE'},
  {key:'freshfood',    label:'FRESH FOOD'},
  {key:'mmpp',         label:'MMPP'},
  {key:'clientes',     label:'CLIENTES META'},
  {key:'activacion',   label:'ACTIVACIÓN'},
  {key:'aa',           label:'A.A. — AÑO ANTERIOR'},
  {key:'margen',       label:'MARGEN'},
  {key:'visitas',      label:'VISITAS DIARIAS'},
  {key:'conversion',   label:'CONVERSIÓN'},
  {key:'productividad',label:'PRODUCTIVIDAD'},
  {key:'clima',        label:'CLIMA MANDELA'},
  {key:'cdo',          label:'CDO PERFORMAS: 100 PUNTOS'}
];

function allTableOptions(includeBase){
  var opts = '<option value="">— Elige una tabla —</option>';
  if(includeBase){
    BASE_TABLES.forEach(function(b){ opts+='<option value="'+b.key+'">'+b.label+'</option>'; });
  }
  Object.keys(customTables).forEach(function(k){
    opts+='<option value="'+k+'">'+customTables[k].name+'</option>';
  });
  return opts;
}

function populateEditSel(){
  var sel=document.getElementById('edit-table-sel'); if(!sel)return;
  sel.innerHTML=allTableOptions(true);
}
function populateAddrowSel(){
  var sel=document.getElementById('addrow-table-sel'); if(!sel)return;
  sel.innerHTML=allTableOptions(true);
}
function populateEditColsSel(){
  var sel=document.getElementById('editcols-sel'); if(!sel)return;
  // Editar columnas aplica a todas las tablas
  sel.innerHTML=allTableOptions(true);
}
function populateAddColSel(){
  var sel=document.getElementById('addcol-sel'); if(!sel)return;
  sel.innerHTML=allTableOptions(true);
}
function populateDeltableSel(){
  var sel=document.getElementById('deltable-sel'); if(!sel)return;
  var opts='<option value="">— Elige una tabla personalizada —</option>';
  Object.keys(customTables).forEach(function(k){
    opts+='<option value="'+k+'">'+customTables[k].name+'</option>';
  });
  sel.innerHTML=opts;
  if(!Object.keys(customTables).length){
    sel.innerHTML='<option value="">Sin tablas personalizadas creadas</option>';
  }
}
function populateDelrowSel(){
  var sel=document.getElementById('delrow-table-sel'); if(!sel)return;
  sel.innerHTML=allTableOptions(true);
}

// ── Editar tabla ──────────────────────────────
// Para tablas tipo avance: Av.Lineal = Avance/Presupuesto, Proy.Cierre = Proyección/Presupuesto
function calcAvanceRow(r){
  var pres = parseFloat(r[1].replace(/,/g,''));
  var av   = parseFloat(r[2].replace(/,/g,''));
  var proy = parseFloat(r[4].replace(/,/g,''));
  if(!isNaN(pres) && pres>0){
    if(!isNaN(av))   r[3] = (av/pres*100).toFixed(1)+'%';
    if(!isNaN(proy)) r[5] = (proy/pres*100).toFixed(1)+'%';
  }
}

function isAvanceTable(key){
  return key==='avance'||key==='freshfood'||key==='mmpp';
}

function loadEditRows(){
  var key=document.getElementById('edit-table-sel').value;
  var con=document.getElementById('edit-rows-container');
  var btn=document.getElementById('btn-save-edit');
  if(!key){con.innerHTML='';btn.style.display='none';return;}
  var rows=getRows(key), cols=getCols(key);
  btn.style.display='';

  // Para tablas de avance, las cols auto-calculadas son: Av.Lineal(3) y Proy.Cierre(5)
  var autoCalcCols = AUTO_COLS[key] || {};

  var html='<div style="overflow-x:auto;margin-top:12px">'
    +'<table style="width:100%;border-collapse:collapse;font-size:12px">'
    +'<thead><tr>'
    +cols.map(function(c,ci){
      var isAuto = autoCalcCols[ci];
      return '<th style="background:'+(isAuto?'#f0f7ff':'#f5f7fb')+';border:1px solid #e2e6f0;'
        +'padding:8px 10px;text-align:left;font-size:10px;font-weight:700;'
        +'color:'+(isAuto?'#2e74e8':'#8490a8')+';text-transform:uppercase;white-space:nowrap">'
        +c+(isAuto?' 🔄':'')+'</th>';
    }).join('')
    +'</tr></thead><tbody>';

  rows.forEach(function(r,ri){
    var isPS = r[0]==='PRESTADOR S.';
    html+='<tr style="'+(ri%2===0?'background:#fff':'background:#f9fafb')+'">';
    cols.forEach(function(c,ci){
      var val = r[ci]!==undefined ? r[ci] : '—';
      var isAuto = autoCalcCols[ci];
      if(isAuto){
        // Columna auto-calculada: solo lectura, fondo azul claro
        html+='<td style="border:1px solid #e2e6f0;padding:4px 6px;background:#f0f7ff;'
          +'text-align:right;font-size:12px;color:#2e74e8;font-weight:600;padding:8px 10px">'
          +val+'</td>';
      } else {
        html+='<td style="border:1px solid #e2e6f0;padding:4px 6px">'
          +'<input class="fi" data-ri="'+ri+'" data-ci="'+ci+'" value="'+val+'"'
          +' style="width:100%;min-width:80px;padding:5px 8px;font-size:12px;border:1px solid transparent;'
          +'border-radius:5px;background:transparent;font-family:Inter,sans-serif"'
          +' onfocus="this.style.border=\'1px solid #2e74e8\';this.style.background=\'#fff\'"'
          +' onblur="this.style.border=\'1px solid transparent\';this.style.background=\'transparent\'">'
          +'</td>';
      }
    });
    html+='</tr>';
  });

  // Note dinámica según tabla
  var autoCNames=Object.keys(autoCalcCols).map(function(ci){return cols[parseInt(ci)]||'';}).filter(Boolean);
  var noteHtml='</tbody></table></div>';
  if(autoCNames.length>0){
    noteHtml+='<div style="margin-top:10px;font-size:11px;color:#2e74e8;background:#f0f7ff;'
      +'padding:9px 14px;border-radius:7px;line-height:1.6">'
      +'🔄 <b>'+autoCNames.join('</b>, <b>')+'</b> se calculan automáticamente — no es necesario editarlos.<br>'
      +'<span style="color:#8490a8">Edita solo los valores base y el sistema recalculará al guardar.</span></div>';
  } else {
    noteHtml+='<div style="margin-top:10px;font-size:11px;color:#8490a8">💡 Haz clic en cualquier celda para editarla.</div>';
  }
  html+=noteHtml;

  con.innerHTML=html;
}

function saveEdit(){
  var key=document.getElementById('edit-table-sel').value; if(!key)return;
  var rows=getRows(key), cols=getCols(key);

  // Leer todos los inputs editables
  document.querySelectorAll('#edit-rows-container input').forEach(function(inp){
    var ri=parseInt(inp.dataset.ri), ci=parseInt(inp.dataset.ci);
    if(!isNaN(ri)&&!isNaN(ci)&&ri<rows.length){
      rows[ri][ci]=inp.value;
    }
  });

  // Auto-calcular según tipo de tabla
  rows.forEach(function(r){
    if(r[0]==='PRESTADOR S.') return;

    if(isAvanceTable(key)){
      // Av.Lineal = Avance / Presupuesto   (col 3)
      // Proy.Cierre = Proyección / Presupuesto  (col 5)
      calcAvanceRow(r);
    } else {
      // Tablas PRO/OBJ: recalcular Variación y Estado
      var p=parseFloat(r[1]), o=parseFloat(r[2]);
      if(!isNaN(p)&&!isNaN(o)){
        if(cols.length>3 && (cols[3]==='Variación'||cols[3]==='Variación %')) r[3]=calcVar(r[1],r[2]);
        if(cols[cols.length-1]==='Estado') r[cols.length-1]=calcEstado(r[1],r[2]);
      }
      // Productividad: estado vs PRO 2025
      if(cols[0]==='Mes'&&cols[1]==='PRO 2026') r[4]=calcProdEstado(r[1],r[2]);
    }
  });

  saveState();
  closeModal('modal-edit');
  renderTables();
  showToast('✅ Tabla actualizada correctamente');
}

// ── Crear tabla ───────────────────────────────
var selectedHex='#2e74e8';
function selectColor(el){ document.querySelectorAll('.color-opt').forEach(function(c){c.classList.remove('sel');}); el.classList.add('sel'); selectedHex=el.dataset.hex; }
function rebuildRowInputs(){
  var c1=document.getElementById('new-col1').value||'Col1';
  var c2=document.getElementById('new-col2').value||'Col2';
  var c3=document.getElementById('new-col3').value||'Col3';
  var c4=document.getElementById('new-col4').value||'Col4';
  var sec=document.getElementById('dyn-rows-section');
  sec.style.display='block';
  var con=document.getElementById('dyn-rows-container');
  var existing=Array.from(con.querySelectorAll('.dyn-row'));
  // Update headers if any
  if(!existing.length) addNewRowInput();
}
function addNewRowInput(){
  var c=[document.getElementById('new-col1').value||'Col1',document.getElementById('new-col2').value||'Col2',document.getElementById('new-col3').value||'Col3',document.getElementById('new-col4').value||'Col4'];
  var div=document.createElement('div'); div.className='dyn-row';
  div.innerHTML=c.map(function(col,i){return'<input class="fi" placeholder="'+col+'">';}).join('')
    +'<button onclick="this.parentNode.remove()" title="Eliminar">✕</button>';
  document.getElementById('dyn-rows-container').appendChild(div);
  document.getElementById('dyn-rows-section').style.display='block';
}

function saveCreate(){
  var name=document.getElementById('new-table-name').value.trim();
  if(!name){showToast('⚠ Escribe el nombre de la tabla');return;}
  var c=[document.getElementById('new-col1').value.trim()||'Ítem',
         document.getElementById('new-col2').value.trim()||'Valor',
         document.getElementById('new-col3').value.trim()||'OBJ',
         document.getElementById('new-col4').value.trim()||'Estado'];
  var rows=[];
  document.querySelectorAll('#dyn-rows-container .dyn-row').forEach(function(div){
    var ins=div.querySelectorAll('input'); var row=[];
    ins.forEach(function(i){row.push(i.value.trim()||'—');});
    if(row[0]!=='—') rows.push(row);
  });
  var key='custom_'+Date.now();
  customTables[key]={name:name,cols:c,rows:rows,color:selectedHex};
  // Add to addrow/edit selects
  ['addrow-table-sel','edit-table-sel','delrow-table-sel'].forEach(function(sid){
    var sel=document.getElementById(sid); if(!sel)return;
    var opt=document.createElement('option'); opt.value=key; opt.textContent=name; sel.appendChild(opt);
  });
  saveState(); closeModal('modal-create'); renderTables();
  showToast('✅ Tabla "'+name+'" creada');
  // Reset form
  document.getElementById('new-table-name').value='';
  document.getElementById('dyn-rows-container').innerHTML='';
  document.getElementById('dyn-rows-section').style.display='none';
}

// ── Agregar fila ──────────────────────────────
function loadAddRowForm(){
  var key=document.getElementById('addrow-table-sel').value;
  var con=document.getElementById('addrow-form');
  var btn=document.getElementById('btn-save-addrow');
  if(!key){con.innerHTML='';btn.style.display='none';return;}
  var cols=getCols(key);
  btn.style.display='';
  con.innerHTML='<div class="form-row" style="grid-template-columns:repeat('+Math.min(cols.length,2)+',1fr)">'
    +cols.slice(0,4).map(function(c,i){
      return'<div class="fg"><label class="fl">'+c+'</label><input class="fi" id="addrow-f'+i+'" placeholder="'+c+'"></div>';
    }).join('')+'</div>';
}

function saveAddRow(){
  var key=document.getElementById('addrow-table-sel').value; if(!key)return;
  var cols=getCols(key); var rows=getRows(key);
  var newRow=cols.map(function(c,i){
    var inp=document.getElementById('addrow-f'+i);
    return inp?inp.value.trim()||'—':'—';
  });

  if(isAvanceTable(key)){
    // Para tablas de avance: calcular Av.Lineal y Proy.Cierre automáticamente
    calcAvanceRow(newRow);
  } else {
    // Para tablas PRO/OBJ: calcular variación y estado
    if(cols.length>3 && (cols[3]==='Variación'||cols[3]==='Variación %')) newRow[3]=calcVar(newRow[1],newRow[2]);
    if(cols.length>4 && cols[4]==='Estado') newRow[4]=calcEstado(newRow[1],newRow[2]);
    // Productividad
    if(cols[0]==='Mes'&&cols[1]==='PRO 2026') newRow[4]=calcProdEstado(newRow[1],newRow[2]);
  }

  rows.push(newRow);
  saveState(); closeModal('modal-addrow'); renderTables();
  showToast('✅ Fila agregada');
}

// ── Editar columnas ───────────────────────────
function loadEditColsForm(){
  var key=document.getElementById('editcols-sel').value;
  var con=document.getElementById('editcols-form');
  var btn=document.getElementById('btn-save-cols');
  if(!key){con.innerHTML='';btn.style.display='none';return;}
  var cols=getCols(key);
  btn.style.display='';
  con.innerHTML='<div class="form-row" style="grid-template-columns:1fr 1fr">'
    +cols.map(function(c,i){
      return'<div class="fg"><label class="fl">Columna '+(i+1)+'</label>'
        +'<input class="fi" id="editcol-'+i+'" value="'+c+'" placeholder="Nombre columna '+(i+1)+'"></div>';
    }).join('')+'</div>';
}

function saveEditCols(){
  var key=document.getElementById('editcols-sel').value; if(!key)return;
  var cols=getCols(key);
  cols.forEach(function(c,i){
    var inp=document.getElementById('editcol-'+i);
    if(inp&&inp.value.trim()) cols[i]=inp.value.trim();
  });
  // Apply back to correct store
  if(tableData[key]) tableData[key].cols=cols;
  else if(customTables[key]) customTables[key].cols=cols;
  saveState(); closeModal('modal-editcols'); renderTables();
  showToast('✅ Columnas actualizadas');
}

// ── Agregar columna ───────────────────────────
function saveAddCol(){
  var key=document.getElementById('addcol-sel').value;
  var colName=document.getElementById('addcol-name').value.trim();
  var defVal=document.getElementById('addcol-default').value.trim()||'—';
  if(!key){showToast('⚠ Selecciona una tabla');return;}
  if(!colName){showToast('⚠ Escribe el nombre de la columna');return;}
  var cols=getCols(key);
  var rows=getRows(key);
  // Agregar nueva columna al final
  cols.push(colName);
  rows.forEach(function(r){ r.push(defVal); });
  // Apply back to correct store
  if(tableData[key]){ tableData[key].cols=cols; tableData[key].rows=rows; }
  else if(customTables[key]){ customTables[key].cols=cols; customTables[key].rows=rows; }
  saveState(); closeModal('modal-addcol'); renderTables();
  showToast('✅ Columna "'+colName+'" agregada a la tabla');
}

// ── Eliminar tabla ────────────────────────────
function saveDelTable(){
  var key=document.getElementById('deltable-sel').value;
  if(!key||!customTables[key]){showToast('⚠ Selecciona una tabla');return;}
  var name=customTables[key].name;
  if(!confirm('¿Eliminar la tabla "'+name+'"? Esta acción no se puede deshacer.'))return;
  delete customTables[key];
  // Remove from selects
  ['addrow-table-sel','edit-table-sel','delrow-table-sel'].forEach(function(sid){
    var sel=document.getElementById(sid); if(!sel)return;
    Array.from(sel.options).forEach(function(o){ if(o.value===key) sel.removeChild(o); });
  });
  saveState(); closeModal('modal-deltable'); renderTables();
  showToast('🗑 Tabla "'+name+'" eliminada');
}

// ── Eliminar fila ─────────────────────────────
function loadDelRowForm(){
  var key=document.getElementById('delrow-table-sel').value;
  var con=document.getElementById('delrow-form');
  var btn=document.getElementById('btn-save-delrow');
  if(!key){con.innerHTML='';btn.style.display='none';return;}
  var rows=getRows(key);
  if(!rows.length){
    con.innerHTML='<p style="color:var(--muted);font-size:13px;padding:10px 0">La tabla no tiene filas.</p>';
    btn.style.display='none'; return;
  }
  btn.style.display='';
  con.innerHTML='<div class="fg"><label class="fl">Paso 2 — Selecciona la fila a eliminar</label>'
    +'<select class="fi" id="delrow-row-sel">'
    +rows.map(function(r,i){
      return'<option value="'+i+'">'+r[0]+'  [PRO: '+r[1]+'  OBJ: '+r[2]+']</option>';
    }).join('')
    +'</select></div>';
}

function saveDelRow(){
  var key=document.getElementById('delrow-table-sel').value;
  var idxSel=document.getElementById('delrow-row-sel');
  if(!key||!idxSel)return;
  var idx=parseInt(idxSel.value);
  var rows=getRows(key);
  var seg=rows[idx][0];
  if(!confirm('¿Eliminar la fila "'+seg+'"?'))return;
  rows.splice(idx,1);
  saveState(); closeModal('modal-delrow'); renderTables();
  showToast('✕ Fila "'+seg+'" eliminada');
}

// ═══════════════════════════════════════════════
// PERSISTENCIA
// ═══════════════════════════════════════════════
function saveState(){
  try{
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      tableData: tableData,
      customTables: customTables
    }));
  }catch(e){ console.warn('saveState error:',e); }
}

function loadState(){
  try{
    var raw=localStorage.getItem(STORAGE_KEY);
    if(!raw) return;
    var s=JSON.parse(raw);
    if(!s) return;

    // Restaurar todas las tablas base
    if(s.tableData){
      var baseKeys=['avance','freshfood','mmpp','clientes','activacion',
                    'aa','margen','visitas','conversion','productividad','clima','cdo'];
      baseKeys.forEach(function(k){
        if(s.tableData[k] && s.tableData[k].rows && tableData[k]){
          tableData[k].rows = s.tableData[k].rows;
          // Restaurar cols también si fueron editados
          if(s.tableData[k].cols) tableData[k].cols = s.tableData[k].cols;
        }
      });
    }

    // Restaurar tablas personalizadas
    if(s.customTables){
      customTables = s.customTables;
    }
  }catch(e){ console.warn('loadState error:',e); }
}

// ═══════════════════════════════════════════════
// TOAST
// ═══════════════════════════════════════════════
function showToast(msg){
  var t=document.getElementById('toast');
  t.textContent=msg; t.classList.add('show');
  setTimeout(function(){t.classList.remove('show');},3000);
}


// ═══════════════════════════════════════════════
// INIT
// ═══════════════════════════════════════════════
(function(){
  // Limpiar estado viejo incompatible (tablas antiguas)
  try{
    var raw=localStorage.getItem(STORAGE_KEY);
    if(raw){
      var s=JSON.parse(raw);
      // Si tiene keys viejas (carteras/ventas) pero no nuevas → limpiar
      if(s.tableData && s.tableData.carteras && !s.tableData.avance){
        localStorage.removeItem(STORAGE_KEY);
        console.log('Estado viejo eliminado, usando datos frescos.');
      }
    }
  }catch(e){}
})();
// Las llamadas a loadState/renderTables/updateSbStats se hacen en showApp()

// ═══════════════════════════════════════════════
// LOGIN
// ═══════════════════════════════════════════════
var USERS = [
  {user:'makro', pass:'makro2026', nombre:'Makro Supermayorista', rol:'admin'}
];
var currentUser = null;

function checkLogin(){
  var stored = sessionStorage.getItem('mkLogin');
  if(stored){
    try{ currentUser=JSON.parse(stored); showApp(); return; }catch(e){}
  }
  document.getElementById('login-screen').style.display='flex';
  document.getElementById('app-wrapper').style.display='none';
}

function doLogin(){
  var u=document.getElementById('l-user').value.trim().toLowerCase();
  var p=document.getElementById('l-pass').value;
  var found=USERS.find(function(x){return x.user===u&&x.pass===p;});
  if(found){
    currentUser=found;
    sessionStorage.setItem('mkLogin',JSON.stringify(found));
    showApp();
  } else {
    document.getElementById('l-err').style.display='block';
    document.getElementById('l-pass').value='';
    document.getElementById('l-pass').focus();
  }
}

function showApp(){
  var ls=document.getElementById('login-screen');
  var aw=document.getElementById('app-wrapper');
  if(ls) ls.style.display='none';
  if(aw){ aw.style.display='flex'; }
  var nu=document.getElementById('nav-user');
  if(nu&&currentUser) nu.textContent=currentUser.nombre;
  // Inicializar app si aún no se hizo
  if(!window._appInited){
    window._appInited=true;
    loadState();
    renderTables();
    updateSbStats();
  }
}

function doLogout(){
  sessionStorage.removeItem('mkLogin');
  currentUser=null;
  document.getElementById('app-wrapper').style.display='none';
  document.getElementById('login-screen').style.display='flex';
  document.getElementById('l-user').value='';
  document.getElementById('l-pass').value='';
  document.getElementById('l-err').style.display='none';
}

window.onload = function(){
  checkLogin();
};