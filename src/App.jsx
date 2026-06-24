import { useState, useRef } from "react";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  PageNumber, LevelFormat, TableOfContents, PageBreak, Footer, Header } from "docx";
import { getColumnDescription, getTableContext } from "./prompts.js";

const C={yellow:"#f7c948",dark:"#0f1117",mid:"#1a1f2e",surface:"#141824",border:"#2d3748",text:"#e2e8f0",muted:"#718096",green:"#68d391",blue:"#63b3ed",red:"#fc8181",orange:"#f6ad55"};

function decodeBuffer(buf){
  const b=new Uint8Array(buf);
  if(b[0]===0xFF&&b[1]===0xFE)return new TextDecoder("utf-16le").decode(buf);
  if(b[0]===0xFE&&b[1]===0xFF)return new TextDecoder("utf-16be").decode(buf);
  let n=0;const c=Math.min(b.length,32);
  for(let i=1;i<c;i+=2){if(b[i]===0x00)n++;}
  if(n>=c/4)return new TextDecoder("utf-16le").decode(buf);
  return new TextDecoder("utf-8").decode(buf);
}
async function loadScript(src){
  if(document.querySelector(`script[src="${src}"]`))return;
  return new Promise((res,rej)=>{const s=document.createElement("script");s.src=src;s.onload=res;s.onerror=rej;document.head.appendChild(s);});
}
async function parseFile(file){
  const R={fileName:file.name,sources:[],tables:[],measures:[],relationships:[],pages:[],warnings:[]};
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js");
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pako/2.1.0/pako.min.js");
  const zip=await window.JSZip.loadAsync(await file.arrayBuffer());
  const connFile=zip.file("Connections");
  if(connFile){try{const raw=decodeBuffer(await connFile.async("arraybuffer")).replace(/^\uFEFF/,"");R.sources=(JSON.parse(raw).Connections||[]).map(c=>({name:c.ConnectionString||"",type:c.ConnectionType||""}));}catch(e){R.warnings.push(`Connections: ${e.message}`);}}
  const modelFile=zip.file("DataModelSchema")||zip.file("DataModel");
  if(modelFile){
    try{
      const buf=await modelFile.async("arraybuffer");
      let raw=decodeBuffer(buf).replace(/^\uFEFF/,"");
      if(!raw.trim().startsWith("{")&&!raw.trim().startsWith("[")){try{raw=window.pako.inflate(new Uint8Array(buf),{to:"string"}).replace(/^\uFEFF/,"");}catch(_){}}
      const parsed=JSON.parse(raw);
      const db=Array.isArray(parsed)?(parsed[0]?.model||parsed[0]):(parsed.model||parsed);
      for(const t of(db.tables||[])){
        if(t.name?.startsWith("DateTableTemplate")||t.name?.startsWith("LocalDateTable"))continue;
        const entry={name:t.name,isHidden:!!t.isHidden,columns:[],measures:[]};
        for(const col of(t.columns||[])){if(col.type==="rowNumber")continue;entry.columns.push({name:col.name,dataType:col.dataType||"?",isHidden:!!col.isHidden,calculated:!!col.expression});}
        for(const m of(t.measures||[])){
          const expr=typeof m.expression==="string"?m.expression:(m.expression!=null?String(m.expression):"");
          const me={name:m.name,expression:expr,formatString:m.formatString||"",table:t.name};
          entry.measures.push(me);R.measures.push(me);
        }
        R.tables.push(entry);
      }
      for(const r of(db.relationships||[])){
        R.relationships.push({
          fromTable:r.fromTable,fromColumn:r.fromColumn,
          toTable:r.toTable,toColumn:r.toColumn,
          from:`${r.fromTable}[${r.fromColumn}]`,
          to:`${r.toTable}[${r.toColumn}]`,
          cardinality:r.crossFilteringBehavior||"SingleDirection",
          active:r.isActive!==false
        });
      }
    }catch(e){R.warnings.push(`DataModelSchema: ${e.message}`);}
  }
  const layoutFile=zip.file("Report/Layout");
  if(layoutFile){try{
    const buf=await layoutFile.async("arraybuffer");let raw=decodeBuffer(buf).replace(/^\uFEFF/,"");
    if(!raw.trim().startsWith("{")){try{raw=window.pako.inflate(new Uint8Array(buf),{to:"string"}).replace(/^\uFEFF/,"");}catch(_){}}
    const layout=JSON.parse(raw);
    for(const s of(layout.sections||[])){
      const page={name:s.displayName||s.name||"Página",visuals:[]};
      for(const vc of(s.visualContainers||[])){try{const cfg=typeof vc.config==="string"?JSON.parse(vc.config):vc.config;page.visuals.push({type:cfg?.singleVisual?.visualType||"visual",title:cfg?.singleVisual?.vcObjects?.title?.[0]?.properties?.text?.expr?.Literal?.Value?.replace(/^'|'$/g,"")||""});}catch(_){}}
      R.pages.push(page);
    }
  }catch(e){R.warnings.push(`Layout: ${e.message}`);}}
  return R;
}

async function analyzeWithGroq(parsed, groqKey, addLog){
  const call=async(prompt,label)=>{
    addLog(`Groq: ${label}...`,"info");
    for(let attempt=0;attempt<3;attempt++){
      const res=await fetch("https://api.groq.com/openai/v1/chat/completions",{method:"POST",headers:{"Content-Type":"application/json","Authorization":`Bearer ${groqKey}`},body:JSON.stringify({model:"llama-3.3-70b-versatile",max_tokens:3000,temperature:0.2,messages:[{role:"user",content:prompt}],response_format:{type:"json_object"}})});
      if(res.status===429){
        const wait=attempt===0?36000:72000;
        addLog(`Rate limit — esperando ${wait/1000}s...`,"warn");
        await new Promise(r=>setTimeout(r,wait));continue;
      }
      if(!res.ok){const err=await res.json().catch(()=>({}));throw new Error(err?.error?.message||`Groq ${res.status}`);}
      addLog(`✓ ${label} completado`,"ok");
      return JSON.parse((await res.json()).choices?.[0]?.message?.content||"{}");
    }
    throw new Error("Rate limit persistente. Espera 1 minuto e intenta de nuevo.");
  };

  // CALL 1: framing + pages
  const r1=await call(`Eres experto en Power BI. Responde SOLO JSON sin markdown.
DATOS: ${JSON.stringify({fileName:parsed.fileName,pages:parsed.pages,tableNames:parsed.tables.map(t=>t.name),measureCount:parsed.measures.length,relCount:parsed.relationships.length})}
CONTEXTO: Reporte Power BI de Seguros del Estado. Mide métricas de implementación y testing de proyectos software. Orígenes: Azure DevOps OData Analytics (dos queries: WorkItems y TestPlans, mismo proyecto "segurosdelestado"), SharePoint Lists, VSTS.
{"titulo":"nombre limpio sin extension","introduccion":"4-5 oraciones. Usa 'la organización'. Menciona fuentes tales como SharePoint, Azure DevOps OData y VSTS. Contexto de negocio, objetivos, audiencia.","resumen_ejecutivo":"3-4 oraciones. Menciona específicamente archivo Power BI y reporte en Power BI Service. Qué métricas mide y su valor.","arquitectura":"3-4 oraciones sobre patrón del modelo, capas staging/dimensiones/hechos, convenciones de nomenclatura detectadas.","fuentes":[{"nombre":"Analytics OData - WorkItems","tipo":"OData / Azure DevOps Analytics","endpoint":"[PENDIENTE]","funcion":"Provee work items del proyecto segurosdelestado: épicas, historias, bugs e incidentes para métricas de seguimiento."},{"nombre":"Analytics OData - TestPlans","tipo":"OData / Azure DevOps Analytics","endpoint":"[PENDIENTE]","funcion":"Provee planes de prueba, casos de prueba y resultados de ejecución desde Azure DevOps."},{"nombre":"SharePoint Lista 1","tipo":"SharePoint List","endpoint":"[PENDIENTE]","funcion":"Lista de SharePoint. Completar con nombre y descripción real."},{"nombre":"SharePoint Lista 2","tipo":"SharePoint List","endpoint":"[PENDIENTE]","funcion":"Lista de SharePoint adicional. Completar con nombre y descripción real."}],"paginas":[{"pagina":"nombre exacto","proposito":"qué mide esta página y qué tipos de visuales contiene"}],"conclusiones":"3-4 oraciones con observaciones, buenas prácticas detectadas y recomendaciones de mejora."}`,
  "Call 1/3 — estructura general");

  await new Promise(r=>setTimeout(r,6000));

  // CALL 2: tables + measures (descriptions only, no expressions)
  const tablesCompact=parsed.tables.map(t=>({n:t.name,cols:t.columns.map(c=>c.name).join("|")}));
  const measuresCompact=parsed.measures.map(m=>({n:m.name,t:m.table,e:m.expression.substring(0,120)}));
  const r2=await call(`Eres experto en Power BI y Azure DevOps. Responde SOLO JSON sin markdown.
TABLAS: ${JSON.stringify(tablesCompact).substring(0,4500)}
MEDIDAS (nombre|tabla|expresion_resumida): ${JSON.stringify(measuresCompact).substring(0,3500)}
CONTEXTO: Proyecto segurosdelestado en Azure DevOps. Tablas OData tienen campos estándar de Analytics API. Tablas con prefijo tbl_ son de hechos, DIM_ son dimensiones, SP_ vienen de SharePoint.
Responde:
{"tablas":[{"nombre":"nombre exacto","tipo":"Hecho/Dimensión/Staging/Parámetro/Calendario","origen":"OData Azure DevOps/SharePoint List/Calculada/Combinada","descripcion":"1-2 oraciones función de la tabla","columnas":[{"nombre":"nombre exacto col","descripcion":"qué representa en el negocio, máximo 15 palabras"}]}],"medidas":[{"nombre":"nombre exacto","descripcion":"1 oración: qué calcula y para qué se usa en el reporte"}]}
IMPORTANTE: incluye TODAS las ${parsed.tables.length} tablas y TODAS las ${parsed.measures.length} medidas.`,
  "Call 2/3 — tablas y medidas");

  await new Promise(r=>setTimeout(r,6000));

  // CALL 3: relationships with full names and purpose
  const relsForAI=parsed.relationships.map(r=>({
    from:`${r.fromTable}[${r.fromColumn}]`,
    to:`${r.toTable}[${r.toColumn}]`,
    card:r.cardinality,active:r.active
  }));
  const r3=await call(`Eres experto en Power BI. Responde SOLO JSON sin markdown.
RELACIONES: ${JSON.stringify(relsForAI).substring(0,5000)}
CONTEXTO: Modelo de métricas de implementación/testing en Azure DevOps para Seguros del Estado.
Para cada relación explica en 1 oración su motivo y uso en el modelo. Usa los nombres completos y reales de tablas y columnas.
{"relaciones":[{"from":"tabla[columna] completo","to":"tabla[columna] completo","cardinalidad":"Uno a muchos/Muchos a uno/Muchos a muchos","activa":true,"motivo":"1 oración: por qué existe y cómo se usa para filtrar o cruzar datos"}]}
IMPORTANTE: incluye las ${parsed.relationships.length} relaciones.`,
  "Call 3/3 — relaciones");

  return {...r1,...r2,...r3};
}

async function buildDOCX(docData,parsed){
  const BLUE="1F4E79",ACCENT="2E75B6",GRAY="595959",LGRAY="F2F2F2";
  const fecha=new Date().toLocaleDateString("es-CO",{year:"numeric",month:"long",day:"numeric"});
  const bd={style:BorderStyle.SINGLE,size:1,color:"CCCCCC"};
  const borders={top:bd,bottom:bd,left:bd,right:bd};
  const h1=t=>new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:480,after:160},border:{bottom:{style:BorderStyle.SINGLE,size:8,color:ACCENT,space:4}},children:[new TextRun({text:t,color:BLUE,bold:true,size:36,font:"Calibri"})]});
  const h2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{before:280,after:120},children:[new TextRun({text:t,color:ACCENT,bold:true,size:28,font:"Calibri"})]});
  const h3=t=>new Paragraph({heading:HeadingLevel.HEADING_3,spacing:{before:160,after:80},children:[new TextRun({text:t,color:GRAY,bold:true,size:24,font:"Calibri"})]});
  const pp=(text,opts={})=>new Paragraph({spacing:{after:160,line:276},indent:opts.indent?{left:480}:undefined,alignment:AlignmentType.JUSTIFIED,children:[new TextRun({text:text||"",size:22,font:"Calibri",color:"000000",bold:opts.bold||false})]});
  const kv=(k,v)=>new Paragraph({spacing:{after:100},indent:{left:480},children:[new TextRun({text:`${k}: `,bold:true,size:22,font:"Calibri",color:ACCENT}),new TextRun({text:v||"—",size:22,font:"Calibri"})]});
  const code=t=>new Paragraph({spacing:{before:80,after:80},indent:{left:480,right:480},shading:{fill:"F5F5F5",type:ShadingType.CLEAR},border:{left:{style:BorderStyle.SINGLE,size:12,color:ACCENT,space:8}},children:[new TextRun({text:(t||"").substring(0,600),size:18,font:"Courier New",color:"333333"})]});
  const sp=()=>new Paragraph({spacing:{after:120},children:[new TextRun("")]});
  const pb=()=>new Paragraph({children:[new PageBreak()]});
  const hCell=(t,w)=>new TableCell({borders,shading:{fill:ACCENT,type:ShadingType.CLEAR},width:{size:w,type:WidthType.DXA},margins:{top:80,bottom:80,left:120,right:120},children:[new Paragraph({children:[new TextRun({text:t,bold:true,color:"FFFFFF",size:20,font:"Calibri"})]})]});
  const dCell=(t,w,i,opts={})=>new TableCell({borders,shading:{fill:i%2===0?"FFFFFF":LGRAY,type:ShadingType.CLEAR},width:{size:w,type:WidthType.DXA},margins:{top:60,bottom:60,left:120,right:120},children:[new Paragraph({alignment:AlignmentType.LEFT,children:[new TextRun({text:String(t||""),size:opts.mono?18:20,font:opts.mono?"Courier New":"Calibri",bold:opts.bold||false})]})]});

  const content=[];

  // COVER
  [...Array(6)].forEach(()=>content.push(sp()));
  content.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:240},children:[new TextRun({text:"DOCUMENTACIÓN TÉCNICA",size:60,bold:true,color:BLUE,font:"Calibri"})]}));
  content.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:140},children:[new TextRun({text:docData.titulo||parsed.fileName.replace(/\.(pbix|pbit)$/,""),size:44,color:ACCENT,font:"Calibri",bold:true})]}));
  content.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:80},children:[new TextRun({text:"Power BI — Modelo de Datos",size:28,color:GRAY,font:"Calibri",italics:true})]}));
  [...Array(2)].forEach(()=>content.push(sp()));
  content.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:60},children:[new TextRun({text:`Fecha: ${fecha}`,size:22,font:"Calibri",color:GRAY})]}));
  content.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:60},children:[new TextRun({text:`Tablas: ${parsed.tables.length}  ·  Medidas: ${parsed.measures.length}  ·  Relaciones: ${parsed.relationships.length}  ·  Páginas: ${parsed.pages.length}`,size:22,font:"Calibri",color:GRAY})]}));
  [...Array(4)].forEach(()=>content.push(sp()));
  content.push(pb());

  // TOC
  content.push(h1("Tabla de Contenidos"));
  content.push(new TableOfContents("Tabla de Contenidos",{hyperlink:true,headingStyleRange:"1-3"}));
  content.push(pb());

  // 1. INTRO
  content.push(h1("1. Introducción"));
  content.push(pp(docData.introduccion||""));
  content.push(sp());

  // 2. RESUMEN
  content.push(h1("2. Resumen Ejecutivo"));
  content.push(pp(docData.resumen_ejecutivo||""));
  content.push(sp());

  // 3. FUENTES
  content.push(h1("3. Fuentes de Datos"));
  content.push(pp("A continuación se detallan los orígenes de datos conectados al modelo. Los campos marcados como [PENDIENTE] deben ser completados con la información específica del entorno."));
  content.push(sp());
  for(const f of(docData.fuentes||[])){
    content.push(h2(f.nombre));
    content.push(kv("Tipo de origen",f.tipo));
    content.push(kv("Endpoint / URL de acceso",f.endpoint||"[PENDIENTE - completar]"));
    content.push(kv("Función en el modelo",f.funcion));
    content.push(sp());
  }

  // 4. ARQUITECTURA
  content.push(h1("4. Arquitectura del Modelo"));
  content.push(pp(docData.arquitectura||""));
  content.push(sp());
  content.push(h2("4.1 Diagrama de Flujo de Datos"));
  const diagRows=[
    ["Capa","Componente","Descripción"],
    ["Origen","Azure DevOps OData (Analytics)","Dos queries al proyecto segurosdelestado: WorkItems y TestPlans"],
    ["Origen","SharePoint Lists","Listas con datos de entrada, parámetros y fechas hábiles"],
    ["Origen","VSTS","Conexión a Visual Studio Team Services"],
    ["Staging / Jerarquía","HierarchyFlat / HierarchyTest","Tablas derivadas de OData que representan jerarquía real de work items"],
    ["Modelo semántico","Tablas tbl_ / DIM_ / SP_","Hechos, dimensiones y tablas SharePoint relacionadas"],
    ["Presentación","Power BI Service","Reporte con páginas publicadas y PowerApp embebida"],
  ];
  content.push(new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[1800,2800,4426],rows:diagRows.map((row,i)=>new TableRow({tableHeader:i===0,children:row.map((cell,j)=>i===0?hCell(cell,[1800,2800,4426][j]):dCell(cell,[1800,2800,4426][j],i,{bold:j===0}))}))}));
  content.push(sp());

  // 5. TABLAS
  content.push(h1("5. Tablas y Columnas"));
  content.push(pp(`El modelo contiene ${parsed.tables.length} tablas en total.`));
  content.push(sp());
  const tablesAI=docData.tablas||[];
  for(const pt of parsed.tables){
    const ai=tablesAI.find(t=>t.nombre===pt.name)||{};
    // Check known context first
    const tableDesc=getTableContext(pt.name,ai.descripcion);
    content.push(h2(pt.name));
    content.push(kv("Tipo",ai.tipo||"—"));
    content.push(kv("Origen",ai.origen||"—"));
    if(tableDesc)content.push(pp(tableDesc,{indent:true}));
    if(pt.columns.length){
      content.push(new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[2200,1200,700,700,4226],rows:[
        new TableRow({tableHeader:true,children:[hCell("Columna",2200),hCell("Tipo de Dato",1200),hCell("Calculada",700),hCell("Oculta",700),hCell("Descripción",4226)]}),
        ...pt.columns.slice(0,40).map((col,i)=>{
          const aiCol=(ai.columnas||[]).find(c=>c.nombre===col.name)||{};
          const colDesc=getColumnDescription(pt.name,col.name,aiCol.descripcion);
          return new TableRow({children:[dCell(col.name,2200,i),dCell(col.dataType,1200,i),dCell(col.calculated?"Sí":"No",700,i),dCell(col.isHidden?"Sí":"No",700,i),dCell(colDesc,4226,i)]});
        })
      ]}));
    }
    content.push(sp());
  }

  // 6. MEDIDAS
  content.push(h1("6. Medidas DAX"));
  content.push(pp(`El modelo contiene ${parsed.measures.length} medidas DAX.`));
  content.push(sp());
  const measuresAI=docData.medidas||[];
  const byTable={};
  for(const m of parsed.measures){
    if(!byTable[m.table])byTable[m.table]=[];
    const ai=measuresAI.find(x=>x.nombre===m.name)||{};
    byTable[m.table].push({...m,descripcion:ai.descripcion||""});
  }
  for(const[tbl,measures]of Object.entries(byTable)){
    content.push(h2(`Tabla: ${tbl}`));
    for(const m of measures){
      content.push(h3(`[${m.name}]`));
      if(m.descripcion)content.push(kv("Descripción",m.descripcion));
      if(m.expression)content.push(code(m.expression));
      content.push(sp());
    }
  }

  // 7. RELACIONES
  content.push(h1("7. Relaciones del Modelo"));
  content.push(pp(`El modelo define ${parsed.relationships.length} relaciones entre tablas.`));
  content.push(sp());
  const relsAI=docData.relaciones||[];
  content.push(new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[3400,1500,800,3326],rows:[
    new TableRow({tableHeader:true,children:[hCell("Relación",3400),hCell("Cardinalidad",1500),hCell("Estado",800),hCell("Motivo y uso",3326)]}),
    ...parsed.relationships.map((r,i)=>{
      // Use full readable names
      const rel=`${r.fromTable}[${r.fromColumn}]\n→ ${r.toTable}[${r.toColumn}]`;
      const aiRel=relsAI.find(x=>{
        try{return x.from&&x.from.includes(r.fromTable)&&x.to&&x.to.includes(r.toTable);}catch{return false;}
      })||{};
      const card=r.cardinality==="BothDirections"?"Muchos a muchos":r.cardinality==="ManyToOne"?"Muchos a uno":"Uno a muchos";
      return new TableRow({children:[
        new TableCell({borders,shading:{fill:i%2===0?"FFFFFF":LGRAY,type:ShadingType.CLEAR},width:{size:3400,type:WidthType.DXA},margins:{top:60,bottom:60,left:120,right:120},children:[
          new Paragraph({children:[new TextRun({text:`${r.fromTable}[${r.fromColumn}]`,size:18,font:"Calibri",bold:true})]}),
          new Paragraph({children:[new TextRun({text:`→ ${r.toTable}[${r.toColumn}]`,size:18,font:"Calibri",color:ACCENT})]}),
        ]}),
        dCell(card,1500,i),
        new TableCell({borders,shading:{fill:i%2===0?"FFFFFF":LGRAY,type:ShadingType.CLEAR},width:{size:800,type:WidthType.DXA},margins:{top:60,bottom:60,left:120,right:120},children:[new Paragraph({children:[new TextRun({text:r.active?"✓ Activa":"✗ Inactiva",size:18,font:"Calibri",color:r.active?"276227":"C00000",bold:true})]})]}),
        dCell(aiRel.motivo||"—",3326,i),
      ]});
    })
  ]}));
  content.push(sp());

  // 8. PÁGINAS
  content.push(h1("8. Páginas del Reporte"));
  content.push(pp(`El reporte contiene ${parsed.pages.length} páginas publicadas en Power BI Service.`));
  content.push(sp());
  const paginasAI=docData.paginas||[];
  for(const pg of parsed.pages){
    const ai=paginasAI.find(p=>p.pagina===pg.name)||{};
    content.push(h2(pg.name));
    if(ai.proposito)content.push(pp(ai.proposito,{indent:true}));
    const visuals=pg.visuals.filter(v=>v.type!=="textbox"&&v.type!=="image");
    if(visuals.length){
      const vtypes=[...new Set(visuals.map(v=>v.type))];
      content.push(kv("Tipos de visuales",vtypes.join(", ")));
      content.push(kv("Total de visuales",String(visuals.length)));
    }
    content.push(sp());
  }

  // 9. CONCLUSIONES
  content.push(h1("9. Conclusiones y Recomendaciones"));
  content.push(pp(docData.conclusiones||""));

  const doc=new Document({
    numbering:{config:[{reference:"bullets",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]}]},
    styles:{
      default:{document:{run:{font:"Calibri",size:22}}},
      paragraphStyles:[
        {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:36,bold:true,font:"Calibri",color:"1F4E79"},paragraph:{spacing:{before:480,after:160},outlineLevel:0}},
        {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:28,bold:true,font:"Calibri",color:"2E75B6"},paragraph:{spacing:{before:280,after:120},outlineLevel:1}},
        {id:"Heading3",name:"Heading 3",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:24,bold:true,font:"Calibri",color:"595959"},paragraph:{spacing:{before:200,after:80},outlineLevel:2}},
      ]
    },
    sections:[{
      properties:{page:{size:{width:11906,height:16838},margin:{top:1440,right:1440,bottom:1440,left:1440}}},
      headers:{default:new Header({children:[new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:4,color:ACCENT,space:4}},alignment:AlignmentType.RIGHT,children:[new TextRun({text:`Documentación Técnica — ${docData.titulo||"Power BI"}`,size:18,font:"Calibri",color:GRAY,italics:true})]})]})},
      footers:{default:new Footer({children:[new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:4,color:ACCENT,space:4}},alignment:AlignmentType.CENTER,children:[new TextRun({text:"Página ",size:18,font:"Calibri",color:GRAY}),new TextRun({children:[PageNumber.CURRENT],size:18,font:"Calibri",color:GRAY}),new TextRun({text:" de ",size:18,font:"Calibri",color:GRAY}),new TextRun({children:[PageNumber.TOTAL_PAGES],size:18,font:"Calibri",color:GRAY})]})]})},
      children:content
    }]
  });
  return await Packer.toBlob(doc);
}

const inp={background:"#0a0d13",border:"1px solid #2d3748",borderRadius:8,color:"#e2e8f0",fontSize:13,padding:"10px 14px",width:"100%",boxSizing:"border-box",fontFamily:"monospace",outline:"none"};

export default function App(){
  const[groqKey,setGroqKey]=useState(localStorage.getItem("groq_key")||"");
  const[showKey,setShowKey]=useState(false);const[files,setFiles]=useState([]);const[logs,setLogs]=useState([]);const[progress,setProgress]=useState(0);const[running,setRunning]=useState(false);const[result,setResult]=useState(null);const[dragging,setDragging]=useState(false);
  const logRef=useRef(null);
  const lc={info:C.blue,ok:C.green,warn:C.orange,error:C.red,section:C.yellow};
  const saveKey=k=>{setGroqKey(k);localStorage.setItem("groq_key",k);};
  const addLog=(msg,type="info")=>{setLogs(p=>[...p,{msg,type,time:new Date().toLocaleTimeString("es-CO")}]);setTimeout(()=>logRef.current?.scrollTo(0,logRef.current.scrollHeight),50);};
  const addFiles=fs=>{const valid=fs.filter(f=>f.name.endsWith(".pbix")||f.name.endsWith(".pbit"));setFiles(p=>{const ex=new Set(p.map(f=>f.name));return[...p,...valid.filter(f=>!ex.has(f.name))];});};
  const run=async()=>{
    if(!files.length||running)return;if(!groqKey.trim()){addLog("Ingresa tu Groq API key.","error");return;}
    setRunning(true);setLogs([]);setProgress(0);setResult(null);
    try{for(const file of files){
      addLog(`Procesando: ${file.name}`,"section");setProgress(5);
      addLog("Extrayendo estructura del archivo...","info");
      const parsed=await parseFile(file);setProgress(15);
      parsed.warnings.forEach(w=>addLog(w,"warn"));
      addLog(`✓ Tablas: ${parsed.tables.length} · Medidas: ${parsed.measures.length} · Relaciones: ${parsed.relationships.length} · Páginas: ${parsed.pages.length}`,"ok");
      const docData=await analyzeWithGroq(parsed,groqKey.trim(),addLog);setProgress(75);
      addLog("Generando documento Word profesional...","info");
      const blob=await buildDOCX(docData,parsed);setProgress(95);
      const url=URL.createObjectURL(blob);
      setProgress(100);addLog("✅ Documento Word generado","ok");
      setResult({url,name:file.name.replace(/\.(pbix|pbit)$/,"_documentacion.docx"),docData,parsed});
    }}catch(e){addLog(`Error: ${e.message}`,"error");console.error(e);}finally{setRunning(false);}
  };
  return(
    <div style={{background:C.dark,minHeight:"100vh",fontFamily:"'Segoe UI',system-ui,sans-serif",color:C.text}}>
      <div style={{background:C.surface,borderBottom:`1px solid ${C.border}`,padding:"14px 24px",display:"flex",alignItems:"center",gap:12}}>
        <span style={{fontSize:22}}>📊</span>
        <div><div style={{fontWeight:700,fontSize:15,color:C.yellow}}>PBIX / PBIT Documenter</div><div style={{fontSize:11,color:C.muted}}>Groq · llama-3.3-70b · Word profesional · v9</div></div>
      </div>
      <div style={{maxWidth:760,margin:"0 auto",padding:"20px"}}>
        <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:10,padding:16,marginBottom:16}}>
          <div style={{fontSize:12,color:C.muted,marginBottom:8,fontWeight:600}}>🔑 GROQ API KEY <a href="https://console.groq.com/keys" target="_blank" rel="noreferrer" style={{color:C.blue,marginLeft:10,fontSize:11}}>→ console.groq.com</a></div>
          <div style={{display:"flex",gap:8}}>
            <input type={showKey?"text":"password"} placeholder="gsk_..." value={groqKey} onChange={e=>saveKey(e.target.value)} style={{...inp,flex:1}}/>
            <button onClick={()=>setShowKey(p=>!p)} style={{background:C.mid,border:`1px solid ${C.border}`,borderRadius:8,color:C.muted,padding:"0 14px",cursor:"pointer"}}>{showKey?"🙈":"👁"}</button>
          </div>
          {groqKey&&<div style={{fontSize:11,color:C.green,marginTop:6}}>✓ Key guardada en este navegador</div>}
        </div>
        <div style={{background:"#0d1220",border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 14px",marginBottom:16,fontSize:12,color:C.muted}}>
          ℹ️ 3 llamadas a Groq con pausa de 6s entre cada una. Si hay rate limit, reintenta automáticamente (espera ~36s). Tiempo total estimado: 1-2 minutos.
        </div>
        <div onDragOver={e=>{e.preventDefault();setDragging(true);}} onDragLeave={()=>setDragging(false)} onDrop={e=>{e.preventDefault();setDragging(false);addFiles([...e.dataTransfer.files]);}} onClick={()=>document.getElementById("fi").click()} style={{border:`2px dashed ${dragging?C.yellow:C.border}`,borderRadius:10,padding:"32px 24px",textAlign:"center",cursor:"pointer",background:dragging?C.mid:C.surface,transition:"all 0.2s"}}>
          <input id="fi" type="file" accept=".pbix,.pbit" multiple style={{display:"none"}} onChange={e=>addFiles([...e.target.files])}/>
          <div style={{fontSize:28,marginBottom:8}}>📁</div>
          <div style={{color:C.muted,fontSize:14}}>Arrastra .pbix / .pbit o haz clic</div>
        </div>
        {files.length>0&&(<div style={{marginTop:14}}>
          {files.map((f,i)=>(<div key={i} style={{background:C.mid,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 14px",display:"flex",alignItems:"center",gap:10,marginBottom:8}}><span>📊</span><div style={{flex:1}}><div style={{fontSize:13,fontWeight:600}}>{f.name}</div><div style={{fontSize:11,color:C.muted}}>{(f.size/1024/1024).toFixed(2)} MB</div></div><button onClick={()=>setFiles(p=>p.filter((_,j)=>j!==i))} style={{background:"none",border:"none",color:C.red,cursor:"pointer",fontSize:16}}>✕</button></div>))}
          <button onClick={run} disabled={running} style={{width:"100%",marginTop:4,padding:13,background:running?"#5a4a10":"linear-gradient(135deg,#f7c948,#f59e0b)",color:"#0f1117",fontWeight:700,fontSize:14,border:"none",borderRadius:10,cursor:running?"not-allowed":"pointer"}}>{running?"⏳ Procesando (1-2 min)...":"⚡ Generar Documentación Word"}</button>
          {(running||progress>0)&&(<div style={{height:4,background:C.border,borderRadius:2,overflow:"hidden",marginTop:10}}><div style={{height:"100%",width:`${progress}%`,background:"linear-gradient(90deg,#f7c948,#f59e0b)",transition:"width 0.4s",borderRadius:2}}/></div>)}
        </div>)}
        {logs.length>0&&(<div ref={logRef} style={{background:"#0a0d13",border:`1px solid ${C.border}`,borderRadius:8,padding:14,marginTop:16,maxHeight:260,overflowY:"auto",fontFamily:"monospace",fontSize:12}}>{logs.map((l,i)=><div key={i} style={{color:lc[l.type]||C.text,padding:"1px 0"}}>[{l.time}] {l.msg}</div>)}</div>)}
        {result&&(<div style={{background:"#0d1f0d",border:"1px solid #276227",borderRadius:10,padding:20,marginTop:16}}>
          <div style={{color:C.green,fontWeight:700,fontSize:15,marginBottom:12}}>✅ Documento Word generado</div>
          <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap"}}>{[["Tablas",result.parsed.tables.length],["Medidas",result.parsed.measures.length],["Relaciones",result.parsed.relationships.length],["Páginas",result.parsed.pages.length]].map(([l,n])=>(<div key={l} style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 16px",textAlign:"center",minWidth:70}}><div style={{color:C.yellow,fontWeight:700,fontSize:20}}>{n}</div><div style={{color:C.muted,fontSize:11}}>{l}</div></div>))}</div>
          <a href={result.url} download={result.name} style={{display:"inline-block",padding:"12px 28px",background:"linear-gradient(135deg,#276227,#38a838)",color:"#fff",border:"none",borderRadius:8,fontWeight:700,fontSize:14,textDecoration:"none"}}>⬇ Descargar {result.name}</a>
        </div>)}
      </div>
    </div>
  );
}
