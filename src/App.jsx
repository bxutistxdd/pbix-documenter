import { useState, useRef } from "react";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  PageNumber, LevelFormat, TableOfContents, PageBreak, Footer, Header } from "docx";

const C = { yellow:"#f7c948",dark:"#0f1117",mid:"#1a1f2e",surface:"#141824",border:"#2d3748",text:"#e2e8f0",muted:"#718096",green:"#68d391",blue:"#63b3ed",red:"#fc8181",orange:"#f6ad55" };

function decodeBuffer(buf) {
  const b = new Uint8Array(buf);
  if (b[0]===0xFF&&b[1]===0xFE) return new TextDecoder("utf-16le").decode(buf);
  if (b[0]===0xFE&&b[1]===0xFF) return new TextDecoder("utf-16be").decode(buf);
  let nullOdds=0; const check=Math.min(b.length,32);
  for(let i=1;i<check;i+=2){if(b[i]===0x00)nullOdds++;}
  if(nullOdds>=check/4) return new TextDecoder("utf-16le").decode(buf);
  return new TextDecoder("utf-8").decode(buf);
}
async function loadScript(src){
  if(document.querySelector(`script[src="${src}"]`))return;
  return new Promise((res,rej)=>{const s=document.createElement("script");s.src=src;s.onload=res;s.onerror=rej;document.head.appendChild(s);});
}
async function parseFile(file){
  const result={fileName:file.name,sources:[],tables:[],measures:[],relationships:[],pages:[],roles:[],warnings:[]};
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js");
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pako/2.1.0/pako.min.js");
  const zip=await window.JSZip.loadAsync(await file.arrayBuffer());
  const connFile=zip.file("Connections");
  if(connFile){try{const raw=decodeBuffer(await connFile.async("arraybuffer")).replace(/^\uFEFF/,"");result.sources=(JSON.parse(raw).Connections||[]).map(c=>({name:c.ConnectionString||"Sin nombre",type:c.ConnectionType||"?"}));}catch(e){result.warnings.push(`Connections: ${e.message}`);}}
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
          entry.measures.push(me);result.measures.push(me);
        }
        result.tables.push(entry);
      }
      for(const r of(db.relationships||[])){result.relationships.push({from:`${r.fromTable}[${r.fromColumn}]`,to:`${r.toTable}[${r.toColumn}]`,cardinality:r.crossFilteringBehavior||"SingleDirection",active:r.isActive!==false});}
      for(const role of(db.roles||[])){result.roles.push({name:role.name,tableFilters:(role.tablePermissions||[]).map(p=>({table:p.name,filter:p.filterExpression||""}))});}
    }catch(e){result.warnings.push(`DataModelSchema: ${e.message}`);}
  }else{result.warnings.push("DataModelSchema no encontrado.");}
  const layoutFile=zip.file("Report/Layout");
  if(layoutFile){try{const buf=await layoutFile.async("arraybuffer");let raw=decodeBuffer(buf).replace(/^\uFEFF/,"");if(!raw.trim().startsWith("{")){try{raw=window.pako.inflate(new Uint8Array(buf),{to:"string"}).replace(/^\uFEFF/,"");}catch(_){}}const layout=JSON.parse(raw);for(const section of(layout.sections||[])){const page={name:section.displayName||section.name||"Página",visuals:[]};for(const vc of(section.visualContainers||[])){try{const cfg=typeof vc.config==="string"?JSON.parse(vc.config):vc.config;page.visuals.push({type:cfg?.singleVisual?.visualType||"visual",title:cfg?.singleVisual?.vcObjects?.title?.[0]?.properties?.text?.expr?.Literal?.Value?.replace(/^'|'$/g,"")||""});}catch(_){}}result.pages.push(page);}}catch(e){result.warnings.push(`Layout: ${e.message}`);}}
  return result;
}

async function analyzeWithGroq(parsed,groqKey){
  const summary={fileName:parsed.fileName,sources:parsed.sources,
    tables:parsed.tables.map(t=>({name:t.name,isHidden:t.isHidden,columns:t.columns.map(c=>({name:c.name,type:c.dataType,calculated:c.calculated})),measures:t.measures.map(m=>({name:m.name,expression:m.expression.substring(0,200)}))})),
    relationships:parsed.relationships,pages:parsed.pages,roles:parsed.roles};
  const prompt=`Eres experto en Power BI. Analiza este modelo y genera documentación técnica en español.
MODELO: ${JSON.stringify(summary,null,2).substring(0,14000)}
Responde SOLO JSON válido sin markdown:
{"titulo":"nombre sin extension","resumen_ejecutivo":"3-4 oraciones.","introduccion":"4-5 oraciones de contexto de negocio, objetivos y audiencia.","secciones":[{"id":"fuentes_datos","titulo":"Fuentes de Datos","descripcion":"párrafo","items":[{"nombre":"...","tipo":"...","descripcion":"..."}]},{"id":"arquitectura","titulo":"Arquitectura del Modelo","descripcion":"párrafo sobre patrón, capas y nomenclatura.","items":[]},{"id":"tablas","titulo":"Tablas y Columnas","descripcion":"párrafo","items":[{"tabla":"...","tipo":"Hecho/Dimensión/Staging/Parámetro","descripcion":"...","columnas_clave":["..."]}]},{"id":"medidas","titulo":"Medidas DAX","descripcion":"párrafo","items":[{"nombre":"...","tabla":"...","expresion":"...","proposito":"...","categoria":"KPI/Auxiliar/Filtro"}]},{"id":"relaciones","titulo":"Relaciones del Modelo","descripcion":"párrafo","items":[]},{"id":"paginas","titulo":"Páginas del Reporte","descripcion":"párrafo","items":[{"pagina":"...","proposito":"...","visuales":["..."]}]}],"conclusiones":"3-4 oraciones con recomendaciones."}`;
  const res=await fetch("https://api.groq.com/openai/v1/chat/completions",{method:"POST",headers:{"Content-Type":"application/json","Authorization":`Bearer ${groqKey}`},body:JSON.stringify({model:"llama-3.3-70b-versatile",max_tokens:5000,temperature:0.2,messages:[{role:"user",content:prompt}],response_format:{type:"json_object"}})});
  if(!res.ok){const err=await res.json().catch(()=>({}));throw new Error(err?.error?.message||`Groq ${res.status}`);}
  return JSON.parse((await res.json()).choices?.[0]?.message?.content||"{}");
}

async function buildDOCX(docData, parsed) {
  const BLUE="1F4E79", ACCENT="2E75B6", GRAY="595959", LIGHT_GRAY="F2F2F2";
  const fecha=new Date().toLocaleDateString("es-CO",{year:"numeric",month:"long",day:"numeric"});
  const border={style:BorderStyle.SINGLE,size:1,color:"CCCCCC"};
  const borders={top:border,bottom:border,left:border,right:border};

  const h1=(text)=>new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:480,after:160},border:{bottom:{style:BorderStyle.SINGLE,size:8,color:ACCENT,space:4}},children:[new TextRun({text,color:BLUE,bold:true,size:36,font:"Calibri"})]});
  const h2=(text)=>new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{before:280,after:120},children:[new TextRun({text,color:ACCENT,bold:true,size:28,font:"Calibri"})]});
  const p=(text,opts={})=>new Paragraph({spacing:{after:160,line:276},indent:opts.indent?{left:480}:undefined,alignment:AlignmentType.JUSTIFIED,children:[new TextRun({text:text||"",size:22,font:"Calibri",color:opts.color||"000000",bold:opts.bold||false,italics:opts.italic||false})]});
  const kv=(key,val)=>new Paragraph({spacing:{after:100},indent:{left:480},children:[new TextRun({text:`${key}: `,bold:true,size:22,font:"Calibri",color:ACCENT}),new TextRun({text:val||"—",size:22,font:"Calibri"})]});
  const code=(text)=>new Paragraph({spacing:{before:80,after:80},indent:{left:480,right:480},shading:{fill:"F5F5F5",type:ShadingType.CLEAR},border:{left:{style:BorderStyle.SINGLE,size:12,color:ACCENT,space:8}},children:[new TextRun({text:text.substring(0,500),size:18,font:"Courier New",color:"333333"})]});
  const sp=()=>new Paragraph({spacing:{after:120},children:[new TextRun("")]});
  const pb=()=>new Paragraph({children:[new PageBreak()]});

  const hdrCell=(text,w)=>new TableCell({borders,shading:{fill:ACCENT,type:ShadingType.CLEAR},width:{size:w,type:WidthType.DXA},margins:{top:80,bottom:80,left:120,right:120},children:[new Paragraph({children:[new TextRun({text,bold:true,color:"FFFFFF",size:20,font:"Calibri"})]})]});
  const dataCell=(text,w,idx,bold=false)=>new TableCell({borders,shading:{fill:idx%2===0?"FFFFFF":LIGHT_GRAY,type:ShadingType.CLEAR},width:{size:w,type:WidthType.DXA},margins:{top:60,bottom:60,left:120,right:120},children:[new Paragraph({children:[new TextRun({text:String(text||""),size:20,font:"Calibri",bold})]})]});

  const sectionNums={fuentes_datos:3,arquitectura:4,tablas:5,medidas:6,relaciones:7,paginas:8};
  const content=[];

  // COVER
  content.push(...Array(6).fill(null).map(sp));
  content.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:200},children:[new TextRun({text:"DOCUMENTACIÓN TÉCNICA",size:60,bold:true,color:BLUE,font:"Calibri"})]}));
  content.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:120},children:[new TextRun({text:docData.titulo||parsed.fileName.replace(/\.(pbix|pbit)$/,""),size:44,color:ACCENT,font:"Calibri",bold:true})]}));
  content.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:80},children:[new TextRun({text:"Power BI — Modelo de Datos",size:28,color:GRAY,font:"Calibri",italics:true})]}));
  content.push(sp(),sp());
  content.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:60},children:[new TextRun({text:`Fecha: ${fecha}`,size:22,font:"Calibri",color:GRAY})]}));
  content.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:60},children:[new TextRun({text:`Tablas: ${parsed.tables.length}  ·  Medidas: ${parsed.measures.length}  ·  Relaciones: ${parsed.relationships.length}  ·  Páginas: ${parsed.pages.length}`,size:22,font:"Calibri",color:GRAY})]}));
  content.push(...Array(4).fill(null).map(sp));
  content.push(pb());

  // TOC
  content.push(h1("Tabla de Contenidos"));
  content.push(new TableOfContents("Tabla de Contenidos",{hyperlink:true,headingStyleRange:"1-3"}));
  content.push(pb());

  // 1. Intro
  content.push(h1("1. Introducción"));
  content.push(p(docData.introduccion||""));
  content.push(sp());

  // 2. Resumen
  content.push(h1("2. Resumen Ejecutivo"));
  content.push(p(docData.resumen_ejecutivo||""));
  content.push(sp());

  // Sections
  for(const sec of(docData.secciones||[])){
    const num=sectionNums[sec.id]||"";
    content.push(h1(`${num}. ${sec.titulo}`));
    if(sec.descripcion) content.push(p(sec.descripcion));
    content.push(sp());
    const items=sec.items||[];

    if(sec.id==="fuentes_datos"){
      for(const item of items){content.push(h2(item.nombre));content.push(kv("Tipo",item.tipo));if(item.descripcion)content.push(p(item.descripcion,{indent:true}));content.push(sp());}
    }
    if(sec.id==="tablas"){
      for(const item of items){
        content.push(h2(item.tabla));
        content.push(kv("Tipo",item.tipo));
        if(item.descripcion)content.push(p(item.descripcion,{indent:true}));
        if(item.columnas_clave?.length)content.push(new Paragraph({spacing:{after:100},indent:{left:480},children:[new TextRun({text:"Columnas clave: ",bold:true,size:22,font:"Calibri",color:ACCENT}),new TextRun({text:item.columnas_clave.join(", "),size:22,font:"Calibri"})]}));
        const cols=parsed.tables.find(t=>t.name===item.tabla)?.columns||[];
        if(cols.length){
          content.push(new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[3200,2000,1913,1913],rows:[
            new TableRow({tableHeader:true,children:[hdrCell("Columna",3200),hdrCell("Tipo de Dato",2000),hdrCell("Calculada",1913),hdrCell("Oculta",1913)]}),
            ...cols.slice(0,25).map((col,idx)=>new TableRow({children:[dataCell(col.name,3200,idx),dataCell(col.dataType,2000,idx),dataCell(col.calculated?"Sí":"No",1913,idx),dataCell(col.isHidden?"Sí":"No",1913,idx)]}))
          ]}));
        }
        content.push(sp());
      }
    }
    if(sec.id==="medidas"){
      for(const item of items){
        content.push(h2(`[${item.nombre}]`));
        content.push(kv("Tabla",item.tabla));
        content.push(kv("Categoría",item.categoria||"—"));
        if(item.proposito)content.push(p(item.proposito,{indent:true}));
        if(item.expresion)content.push(code(item.expresion));
        content.push(sp());
      }
    }
    if(sec.id==="relaciones"){
      if(parsed.relationships.length){
        content.push(new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[3400,3400,1500,726],rows:[
          new TableRow({tableHeader:true,children:[hdrCell("Desde",3400),hdrCell("Hasta",3400),hdrCell("Cardinalidad",1500),hdrCell("Estado",726)]}),
          ...parsed.relationships.slice(0,60).map((r,idx)=>new TableRow({children:[dataCell(r.from,3400,idx),dataCell(r.to,3400,idx),dataCell(r.cardinality,1500,idx),new TableCell({borders,shading:{fill:idx%2===0?"FFFFFF":LIGHT_GRAY,type:ShadingType.CLEAR},width:{size:726,type:WidthType.DXA},margins:{top:60,bottom:60,left:120,right:120},children:[new Paragraph({children:[new TextRun({text:r.active?"✓":"✗",size:20,font:"Calibri",color:r.active?"276227":"C00000",bold:true})]})]})]}))
        ]}));
        content.push(sp());
      }
    }
    if(sec.id==="paginas"){
      for(const item of items){content.push(h2(item.pagina));if(item.proposito)content.push(p(item.proposito,{indent:true}));if(item.visuales?.length)content.push(kv("Visuales",item.visuales.join(", ")));content.push(sp());}
    }
  }

  // 9. Conclusions
  content.push(h1("9. Conclusiones y Recomendaciones"));
  content.push(p(docData.conclusiones||""));

  const doc=new Document({
    numbering:{config:[{reference:"bullets",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]}]},
    styles:{
      default:{document:{run:{font:"Calibri",size:22}}},
      paragraphStyles:[
        {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:36,bold:true,font:"Calibri",color:BLUE},paragraph:{spacing:{before:480,after:160},outlineLevel:0}},
        {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:28,bold:true,font:"Calibri",color:ACCENT},paragraph:{spacing:{before:280,after:120},outlineLevel:1}},
        {id:"Heading3",name:"Heading 3",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:24,bold:true,font:"Calibri",color:GRAY},paragraph:{spacing:{before:200,after:80},outlineLevel:2}},
      ]
    },
    sections:[{
      properties:{page:{size:{width:11906,height:16838},margin:{top:1440,right:1440,bottom:1440,left:1440}}},
      headers:{default:new Header({children:[new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:4,color:ACCENT,space:4}},alignment:AlignmentType.RIGHT,children:[new TextRun({text:`Documentación Técnica — ${docData.titulo||"Power BI"}`,size:18,font:"Calibri",color:GRAY,italics:true})]})]})},
      footers:{default:new Footer({children:[new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:4,color:ACCENT,space:4}},alignment:AlignmentType.CENTER,children:[new TextRun({text:"Página ",size:18,font:"Calibri",color:GRAY}),new TextRun({children:[PageNumber.CURRENT],size:18,font:"Calibri",color:GRAY}),new TextRun({text:" de ",size:18,font:"Calibri",color:GRAY}),new TextRun({children:[PageNumber.TOTAL_PAGES],size:18,font:"Calibri",color:GRAY})]})]})},
      children:content
    }]
  });
  return await Packer.toBuffer(doc);
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
      addLog("Extrayendo estructura...","info");
      const parsed=await parseFile(file);setProgress(30);
      parsed.warnings.forEach(w=>addLog(w,"warn"));
      addLog(`✓ Tablas: ${parsed.tables.length} · Medidas: ${parsed.measures.length} · Relaciones: ${parsed.relationships.length} · Páginas: ${parsed.pages.length}`,"ok");
      addLog("Analizando con Groq llama-3.3-70b...","info");
      const docData=await analyzeWithGroq(parsed,groqKey.trim());setProgress(70);addLog("✓ Análisis IA completado","ok");
      addLog("Generando documento Word...","info");
      const buffer=await buildDOCX(docData,parsed);setProgress(95);
      const blob=new Blob([buffer],{type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
      const url=URL.createObjectURL(blob);
      setProgress(100);addLog("✅ Documento Word generado","ok");
      setResult({url,name:file.name.replace(/\.(pbix|pbit)$/,"_documentacion.docx"),docData,parsed});
    }}catch(e){addLog(`Error: ${e.message}`,"error");console.error(e);}finally{setRunning(false);}
  };
  return(
    <div style={{background:C.dark,minHeight:"100vh",fontFamily:"'Segoe UI',system-ui,sans-serif",color:C.text}}>
      <div style={{background:C.surface,borderBottom:`1px solid ${C.border}`,padding:"14px 24px",display:"flex",alignItems:"center",gap:12}}>
        <span style={{fontSize:22}}>📊</span>
        <div><div style={{fontWeight:700,fontSize:15,color:C.yellow}}>PBIX / PBIT Documenter</div><div style={{fontSize:11,color:C.muted}}>Groq · llama-3.3-70b · Exporta Word profesional</div></div>
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
        <div onDragOver={e=>{e.preventDefault();setDragging(true);}} onDragLeave={()=>setDragging(false)} onDrop={e=>{e.preventDefault();setDragging(false);addFiles([...e.dataTransfer.files]);}} onClick={()=>document.getElementById("fi").click()} style={{border:`2px dashed ${dragging?C.yellow:C.border}`,borderRadius:10,padding:"32px 24px",textAlign:"center",cursor:"pointer",background:dragging?C.mid:C.surface,transition:"all 0.2s"}}>
          <input id="fi" type="file" accept=".pbix,.pbit" multiple style={{display:"none"}} onChange={e=>addFiles([...e.target.files])}/>
          <div style={{fontSize:28,marginBottom:8}}>📁</div>
          <div style={{color:C.muted,fontSize:14}}>Arrastra .pbix / .pbit o haz clic</div>
        </div>
        {files.length>0&&(<div style={{marginTop:14}}>
          {files.map((f,i)=>(<div key={i} style={{background:C.mid,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 14px",display:"flex",alignItems:"center",gap:10,marginBottom:8}}><span>📊</span><div style={{flex:1}}><div style={{fontSize:13,fontWeight:600}}>{f.name}</div><div style={{fontSize:11,color:C.muted}}>{(f.size/1024/1024).toFixed(2)} MB</div></div><button onClick={()=>setFiles(p=>p.filter((_,j)=>j!==i))} style={{background:"none",border:"none",color:C.red,cursor:"pointer",fontSize:16}}>✕</button></div>))}
          <button onClick={run} disabled={running} style={{width:"100%",marginTop:4,padding:13,background:running?"#5a4a10":"linear-gradient(135deg,#f7c948,#f59e0b)",color:"#0f1117",fontWeight:700,fontSize:14,border:"none",borderRadius:10,cursor:running?"not-allowed":"pointer"}}>{running?"⏳ Procesando...":"⚡ Generar Documentación Word"}</button>
          {(running||progress>0)&&(<div style={{height:4,background:C.border,borderRadius:2,overflow:"hidden",marginTop:10}}><div style={{height:"100%",width:`${progress}%`,background:"linear-gradient(90deg,#f7c948,#f59e0b)",transition:"width 0.4s",borderRadius:2}}/></div>)}
        </div>)}
        {logs.length>0&&(<div ref={logRef} style={{background:"#0a0d13",border:`1px solid ${C.border}`,borderRadius:8,padding:14,marginTop:16,maxHeight:220,overflowY:"auto",fontFamily:"monospace",fontSize:12}}>{logs.map((l,i)=><div key={i} style={{color:lc[l.type]||C.text,padding:"1px 0"}}>[{l.time}] {l.msg}</div>)}</div>)}
        {result&&(<div style={{background:"#0d1f0d",border:"1px solid #276227",borderRadius:10,padding:20,marginTop:16}}>
          <div style={{color:C.green,fontWeight:700,fontSize:15,marginBottom:12}}>✅ Documento Word generado</div>
          <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap"}}>{[["Tablas",result.parsed.tables.length],["Medidas",result.parsed.measures.length],["Relaciones",result.parsed.relationships.length],["Páginas",result.parsed.pages.length]].map(([l,n])=>(<div key={l} style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 16px",textAlign:"center",minWidth:70}}><div style={{color:C.yellow,fontWeight:700,fontSize:20}}>{n}</div><div style={{color:C.muted,fontSize:11}}>{l}</div></div>))}</div>
          <a href={result.url} download={result.name} style={{display:"inline-block",padding:"12px 28px",background:"linear-gradient(135deg,#276227,#38a838)",color:"#ffffff",border:"none",borderRadius:8,fontWeight:700,fontSize:14,textDecoration:"none"}}>⬇ Descargar {result.name}</a>
        </div>)}
      </div>
    </div>
  );
}
