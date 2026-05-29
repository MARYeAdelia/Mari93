import { useState, useCallback } from "react";
import * as XLSX from "xlsx";

const ANALISTAS = ["MARIANA","WILDER","GIOVANNI"];
const LIDERANCA = ["CARLA","DARLAN"];
const COLORS = { MARIANA:"#E8A87C",WILDER:"#7CB9E8",GIOVANNI:"#A8E87C",CARLA:"#C8A96E",DARLAN:"#B87CE8" };
const CAT_COM = ["Reajuste","Renovação","Up Selling","Defesa de Território","Alteração de Escopo","Outros"];
const CAT_COM_C = { "Reajuste":"#E8A87C","Renovação":"#7CB9E8","Up Selling":"#A8E87C","Defesa de Território":"#C8A96E","Alteração de Escopo":"#B87CE8","Outros":"#5A5A5A" };
const CAT_ENT_C = { "PEC":"#7CB9E8","Abertura de Custo":"#A8E87C","Proposta Comercial":"#C8A96E","Carta de Reajuste":"#E8A87C","Notificação de Reajuste":"#B87CE8","Revisão de Escopo":"#E87C9A" };
const SIN_C  = { verde:"#3A8A3A",amarelo:"#8A8A2A",vermelho:"#8A2A2A" };
const SIN_BG = { verde:"#0A2A0A",amarelo:"#2A2A08",vermelho:"#2A0A0A" };
const SIN_LB = { verde:"🟢 Boa Negociação",amarelo:"🟡 Moderada",vermelho:"🔴 Difícil" };

const fmt = (v) => (!v&&v!==0)?"—":new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL",maximumFractionDigits:0}).format(v);
const fmtPct = (v) => (v!=null&&v!==""&&!isNaN(v))?`${(parseFloat(v)*100).toFixed(1)}%`:"—";
const n = (v) => parseFloat(v)||0;

const matchPessoa = (farmer,hunter) => {
  const chk = (s) => {
    s=(s||"").toUpperCase();
    if(s.includes("MARIANA")) return "MARIANA";
    if(s.includes("WILDER"))  return "WILDER";
    if(s.includes("GIOVANNI"))return "GIOVANNI";
    if(s.includes("CARLA"))   return "CARLA";
    if(s.includes("DARLAN"))  return "DARLAN";
    return null;
  };
  return chk(farmer)||chk(hunter)||null;
};

const matchCat = (ativ,tipo) => {
  const c=((ativ||"")+" "+(tipo||"")).toUpperCase();
  if(c.includes("REAJUSTE")||c.includes("NOTIFICAÇ")||c.includes("CARTA DE REAJUSTE")) return "Reajuste";
  if(c.includes("RENOVAÇ")) return "Renovação";
  if(c.includes("UP-SELLING")||c.includes("UP SELLING")||c.includes("AUMENTO DE ESCOPO")||c.includes("PROPOSTA AUMENTO")) return "Up Selling";
  if(c.includes("DEFESA DE TERRIT")) return "Defesa de Território";
  if((c.includes("ALTERAÇ")&&c.includes("ESCOPO"))||c.includes("REVISÃO DE ESCOPO")||c.includes("REVISAO DE ESCOPO")||c.includes("ADITIVO CONTRATUAL")) return "Alteração de Escopo";
  return "Outros";
};

const matchEnt = (r) => {
  const found=new Set();
  const isSim=(v)=>{const s=(v||"").toString().toUpperCase().trim();return s==="SIM"||s==="S";};
  if(isSim(r.fezPec))      found.add("PEC");
  if(isSim(r.fezAbertura)) found.add("Abertura de Custo");
  if(isSim(r.fezProposta)) found.add("Proposta Comercial");
  if(isSim(r.fezCarta))    found.add("Carta de Reajuste");
  if(isSim(r.fezNotif))    found.add("Notificação de Reajuste");
  if(!found.size){
    const a=(r.atividade||"").toUpperCase(),o=(r.obs||"").toUpperCase();
    if(a.includes("PEC")||o.includes("PEC")) found.add("PEC");
    if(a.includes("OPEN BOOK")||a.includes("ABERTURA")||o.includes("ABERTURA DE CUSTO")) found.add("Abertura de Custo");
    if(a.includes("PROPOSTA COMERCIAL")||o.includes("PROPOSTA COMERCIAL")) found.add("Proposta Comercial");
    if(a.includes("CARTA DE REAJUSTE")||o.includes("CARTA DE REAJUSTE")) found.add("Carta de Reajuste");
    if(a.includes("NOTIFICAÇ")||o.includes("NOTIFICAÇ")) found.add("Notificação de Reajuste");
    if(a.includes("REVISÃO DE ESCOPO")||o.includes("REVISÃO DE ESCOPO")) found.add("Revisão de Escopo");
  }
  return found.size?[...found]:[];
};

const matchSin = (v) => {
  const s=(v||"").toString().toUpperCase();
  if(s.includes("VERDE")) return "verde";
  if(s.includes("AMARELO")) return "amarelo";
  if(s.includes("VERM")) return "vermelho";
  return null;
};

const parseExcel = (file) => new Promise((res,rej)=>{
  const rd=new FileReader();
  rd.onload=(e)=>{
    try{
      const wb=XLSX.read(e.target.result,{type:"array"});
      const nm=wb.SheetNames.find(s=>s.toUpperCase()==="DADOS")||wb.SheetNames[0];
      res(XLSX.utils.sheet_to_json(wb.Sheets[nm],{defval:""}));
    }catch(err){rej(err);}
  };
  rd.onerror=rej;
  rd.readAsArrayBuffer(file);
});

const processRows = (rows) => {
  if(!rows||!rows.length) return [];
  const keys = Object.keys(rows[0]);
  const col = (...kws) => keys.find(k => kws.some(kw => k.toUpperCase().replace(/\s+/g," ").trim().includes(kw.toUpperCase()))) || null;
  const val = (r, ...kws) => { const k=col(...kws); return k ? r[k] : ""; };
  console.log("Colunas:", keys);
  return rows.map(r=>{
    const ativ=(val(r,"ATIVIDADE1")||val(r,"ATIVIDADE")||"").toString().trim().toUpperCase();
    const tipo=(val(r,"TIPO DE PROPOSTA")||"").toString().trim().toUpperCase();
    const obs=(val(r,"OBS")||"").toString().trim();
    const cValAtual = col("VALOR CONTRATO ATUAL");
    const valAtual = cValAtual ? n(r[cValAtual]) : 0;
    const cValPleito = col("REAJUSTE + PLEITO","COM REAJUSTE + PLEITO");
    const cValReaj   = col("VALOR CONTRATO COM REAJUSTE");
    const valPleito  = (cValPleito ? n(r[cValPleito]) : 0) || (cValReaj ? n(r[cValReaj]) : 0);
    const row={
      isRevisao:(val(r,"REVISÃO","REVISAO")||"").toString().toUpperCase().includes("REVIS"),
      nProposta:(val(r,"Nº PROPOSTA","N PROPOSTA","PROPOSTA")||"").toString().trim(),
      grupoCliente:(val(r,"GRUPO CLIENTE")||"").toString().trim(),
      cliente:(val(r,"CLIENTE")||"").toString().trim(),
      respFarmer:(val(r,"RESP. FARMER","RESPONSÁVEL FARMER")||"").toString().trim(),
      respHunter:(val(r,"RESP. HUNTER","RESPONSÁVEL HUNTER")||"").toString().trim(),
      atividade:ativ, tipoProposta:tipo, obs,
      valAtual, valPleito, diferenca: valPleito - valAtual,
      pctReaj:   n(val(r,"REAJUSTE CONTRATUAL (%)")),
      pctPleito: n(val(r,"PLEITO (%)")),
      aprovR:    n(val(r,"APROVADO PELO CLIENTE (R$)")),
      aprovPct:  n(val(r,"APROVADO PELO CLIENTE (%)")),
      status:    (val(r,"STATUS")||"").toString().trim(),
      mes:       (val(r,"Mês","MES")||"").toString().trim(),
      sinalizacao: matchSin(val(r,"Sinalização","SINALIZACAO","SINALIZ")||""),
      fezPec:      val(r,"FEZ PEC"),
      fezAbertura: val(r,"FEZ ABERTURA"),
      fezProposta: val(r,"FEZ PROPOSTA COMERCIAL"),
      fezCarta:    val(r,"FEZ CARTA DE REAJUSTE"),
      fezNotif:    val(r,"FEZ CARTA DE NOTIF","FEZ NOTIF"),
    };
    row.responsavel=matchPessoa(row.respFarmer,row.respHunter);
    row.categoria=matchCat(ativ,tipo);
    row.entregaveis=matchEnt(row);
    return row;
  }).filter(r=>r.responsavel);
};

const Bar = ({p,color}) => (
  <div style={{flex:1,height:5,background:"#1A1A22",borderRadius:3,overflow:"hidden"}}>
    <div style={{width:`${Math.min(p,100)}%`,height:"100%",background:color,borderRadius:3,transition:"width .4s"}}/>
  </div>
);
const Tag = ({label,color}) => (
  <span style={{fontSize:9,padding:"2px 7px",background:`${color}22`,color,border:`1px solid ${color}44`,whiteSpace:"nowrap",borderRadius:2}}>{label}</span>
);
const Btn = ({label,active,color="#C8A96E",onClick}) => (
  <button onClick={onClick} style={{padding:"5px 13px",border:"1px solid",cursor:"pointer",fontFamily:"inherit",borderColor:active?color:"#2A2A2E",background:active?`${color}18`:"transparent",color:active?color:"#5A5A5A",fontSize:10,letterSpacing:2,textTransform:"uppercase"}}>{label}</button>
);
const Mini = ({label,val,highlight}) => (
  <div style={{background:"#0D0D0F",padding:"7px 10px"}}>
    <div style={{fontSize:8,letterSpacing:2,color:"#4A4A4A",textTransform:"uppercase",marginBottom:3}}>{label}</div>
    <div style={{fontSize:11,color:highlight?"#C8A96E":"#F0EDE8"}}>{val||"—"}</div>
  </div>
);
const SecLabel = ({children}) => (
  <div style={{fontSize:9,letterSpacing:4,color:"#5A5A5A",textTransform:"uppercase",marginBottom:10}}>{children}</div>
);

export default function FarmDashboard() {
  const [data,setData]=useState([]);
  const [view,setView]=useState("visao");
  const [fResp,setFResp]=useState("Todos");
  const [fStat,setFStat]=useState("Todos");
  const [fCat,setFCat]=useState("Todas");
  const [fSin,setFSin]=useState("Todos");
  const [fileName,setFN]=useState(null);
  const [loading,setLd]=useState(false);
  const [drag,setDrag]=useState(false);
  const [error,setErr]=useState(null);

  const load=useCallback(async(file)=>{
    setLd(true);setErr(null);
    try{setData(processRows(await parseExcel(file)));setFN(file.name);}
    catch{setErr("Erro ao ler o arquivo. Verifique se é um .xlsx válido.");}
    finally{setLd(false);}
  },[]);

  const ativs=data.filter(r=>!r.isRevisao);
  const revs=data.filter(r=>r.isRevisao);
  const analistasAtivs=ativs.filter(r=>ANALISTAS.includes(r.responsavel));

  const applyStatus=(rows)=>{
    if(fStat==="Aprovado")      return rows.filter(r=>r.status.toUpperCase().includes("APROVADO"));
    if(fStat==="Em Negociação") return rows.filter(r=>r.status.toUpperCase().includes("NEGOCI"));
    return rows;
  };
  const filterRows=(rows)=>{
    let r=rows;
    if(fResp!=="Todos") r=r.filter(x=>x.responsavel===fResp);
    r=applyStatus(r);
    if(fCat!=="Todas") r=r.filter(x=>x.categoria===fCat);
    return r;
  };
  const clienteStats=()=>{
    const by={};
    for(const r of ativs){
      const g=r.grupoCliente||r.cliente;
      if(!by[g]) by[g]={rows:[],sin:null};
      by[g].rows.push(r);
      if(!by[g].sin&&r.sinalizacao) by[g].sin=r.sinalizacao;
    }
    return Object.entries(by).map(([g,{rows,sin}])=>{
      const wR=rows.filter(r=>r.pctReaj>0),wA=rows.filter(r=>r.aprovPct>0);
      const avg=(arr,fn)=>arr.length?arr.reduce((s,r)=>s+fn(r),0)/arr.length:null;
      return{grupo:g,rows,sin,avgReaj:avg(wR,r=>r.pctReaj),avgPleito:avg(wR,r=>r.pctPleito),avgAprov:avg(wA,r=>r.aprovPct),totalAtual:rows.reduce((s,r)=>s+r.valAtual,0),totalPleito:rows.reduce((s,r)=>s+r.valPleito,0),totalDif:rows.reduce((s,r)=>s+r.diferenca,0)};
    }).sort((a,b)=>b.rows.length-a.rows.length);
  };
  const contarEnt=(rows)=>{
    const c={};let t=0;
    for(const r of rows) for(const e of r.entregaveis){c[e]=(c[e]||0)+1;t++;}
    return Object.entries(c).map(([cat,qtd])=>({cat,qtd,p:t?Math.round((qtd/t)*100):0})).sort((a,b)=>b.qtd-a.qtd);
  };

  const filtHistorico=filterRows(data);

  return (
    <div style={{minHeight:"100vh",background:"#0D0D0F",color:"#F0EDE8",fontFamily:"'Georgia',serif"}}>
      <div style={{borderBottom:"1px solid #1E1E24",padding:"16px 28px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
        <div>
          <div style={{fontSize:10,letterSpacing:4,color:"#5A5A5A",textTransform:"uppercase",marginBottom:2}}>Time Comercial</div>
          <div style={{fontSize:20}}>Farm <span style={{color:"#C8A96E"}}>Performance</span></div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          {fileName&&<div style={{fontSize:10,color:"#4A4A4A",background:"#13131A",border:"1px solid #1E1E24",padding:"4px 10px"}}>📄 {fileName}</div>}
          <label style={{cursor:"pointer"}}>
            <input type="file" accept=".xlsx,.xls" onChange={e=>{const f=e.target.files[0];if(f)load(f);}} style={{display:"none"}}/>
            <div style={{padding:"6px 16px",border:"1px solid #C8A96E",color:"#C8A96E",fontSize:10,letterSpacing:2,textTransform:"uppercase",cursor:"pointer"}}>
              {fileName?"Trocar Planilha":"Carregar Planilha"}
            </div>
          </label>
          {[["visao","Visão Geral"],["tempo","Esforço"],["clientes","Por Cliente"],["historico","Histórico"]].map(([v,l])=>(
            <Btn key={v} label={l} active={view===v} onClick={()=>setView(v)}/>
          ))}
        </div>
      </div>

      <div style={{padding:"22px 28px"}}>
        {!fileName&&(
          <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
            onDrop={e=>{e.preventDefault();setDrag(false);const f=e.dataTransfer.files[0];if(f)load(f);}}
            style={{border:`2px dashed ${drag?"#C8A96E":"#2A2A2E"}`,background:drag?"#C8A96E08":"#13131A",padding:"70px 40px",textAlign:"center",cursor:"pointer"}}>
            <label style={{cursor:"pointer"}}>
              <input type="file" accept=".xlsx,.xls" onChange={e=>{const f=e.target.files[0];if(f)load(f);}} style={{display:"none"}}/>
              <div style={{fontSize:34,marginBottom:12}}>📊</div>
              <div style={{fontSize:14,color:"#C8A96E",marginBottom:6}}>Arraste sua planilha Excel aqui</div>
              <div style={{fontSize:11,color:"#4A4A4A"}}>ou clique em "Carregar Planilha" — aba "Dados" será lida automaticamente</div>
            </label>
          </div>
        )}
        {loading&&<div style={{textAlign:"center",padding:"60px 0",color:"#C8A96E"}}>Lendo planilha...</div>}
        {error&&<div style={{background:"#2A1010",border:"1px solid #5A2020",padding:"12px 18px",color:"#E87C7C",fontSize:13}}>{error}</div>}

        {/* VISÃO GERAL */}
        {fileName&&!loading&&view==="visao"&&(()=>{
          const scopeAll=applyStatus(ativs);
          const scopeCat=(cat)=>scopeAll.filter(r=>r.categoria===cat);
          const cStats=clienteStats();
          return(
            <div>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:18,flexWrap:"wrap"}}>
                <span style={{fontSize:10,letterSpacing:2,color:"#5A5A5A",textTransform:"uppercase"}}>Status:</span>
                {["Todos","Aprovado","Em Negociação"].map(s=><Btn key={s} label={s} active={fStat===s} onClick={()=>setFStat(s)}/>)}
              </div>
              <SecLabel>Totais do Time</SecLabel>
              <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12,marginBottom:28}}>
                {[
                  {l:"Atividades",v:scopeAll.length},{l:"Revisões",v:revs.length},
                  {l:"Valor Atual",v:fmt(scopeAll.reduce((s,r)=>s+r.valAtual,0))},
                  {l:"Valor c/ Pleito",v:fmt(scopeAll.reduce((s,r)=>s+r.valPleito,0))},
                  {l:"Diferença (Ganho)",v:fmt(scopeAll.reduce((s,r)=>s+r.diferenca,0))},
                ].map(i=>(
                  <div key={i.l} style={{background:"#13131A",border:"1px solid #1E1E24",padding:"14px 18px"}}>
                    <div style={{fontSize:9,letterSpacing:3,color:"#5A5A5A",textTransform:"uppercase",marginBottom:4}}>{i.l}</div>
                    <div style={{fontSize:20,color:"#C8A96E"}}>{i.v}</div>
                  </div>
                ))}
              </div>
              <SecLabel>Analistas</SecLabel>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16,marginBottom:32}}>
                {ANALISTAS.map(p=>{
                  const sp=applyStatus(ativs.filter(r=>r.responsavel===p));
                  const revP=revs.filter(r=>r.responsavel===p);
                  const aprov=ativs.filter(r=>r.responsavel===p&&r.status.toUpperCase().includes("APROVADO"));
                  return(
                    <div key={p} style={{background:"#13131A",border:`1px solid ${COLORS[p]}33`,padding:20}}>
                      <div style={{display:"flex",justifyContent:"space-between",marginBottom:14}}>
                        <div>
                          <div style={{fontSize:9,letterSpacing:3,color:"#5A5A5A",textTransform:"uppercase",marginBottom:2}}>Analista</div>
                          <div style={{fontSize:17,color:COLORS[p]}}>{p.charAt(0)+p.slice(1).toLowerCase()}</div>
                        </div>
                        <div style={{textAlign:"right",fontSize:10}}>
                          <div style={{color:"#4A8A4A"}}>✓ {aprov.length} aprovados</div>
                          <div style={{color:"#5A5A3A",marginTop:2}}>{revP.length} revisões</div>
                        </div>
                      </div>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:14}}>
                        <Mini label="Contrato Atual" val={fmt(sp.reduce((s,r)=>s+r.valAtual,0))}/>
                        <Mini label="Com Pleito" val={fmt(sp.reduce((s,r)=>s+r.valPleito,0))}/>
                        <Mini label="Diferença" val={fmt(sp.reduce((s,r)=>s+r.diferenca,0))} highlight/>
                      </div>
                      <div style={{borderTop:"1px solid #1A1A20",paddingTop:12}}>
                        <div style={{fontSize:9,letterSpacing:2,color:"#4A4A4A",textTransform:"uppercase",marginBottom:8}}>Por Tipo</div>
                        {CAT_COM.map(cat=>{
                          const d=sp.filter(r=>r.categoria===cat);
                          if(!d.length) return null;
                          return(
                            <div key={cat} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                              <div style={{display:"flex",alignItems:"center",gap:6}}>
                                <div style={{width:7,height:7,borderRadius:"50%",background:CAT_COM_C[cat]}}/>
                                <span style={{fontSize:11,color:"#8A8A8A"}}>{cat}</span>
                              </div>
                              <div style={{display:"flex",gap:10}}>
                                <span style={{fontSize:11,color:COLORS[p]}}>{d.length}x</span>
                                <span style={{fontSize:11,color:"#C8A96E",minWidth:65,textAlign:"right"}}>{d.reduce((s,r)=>s+r.diferenca,0)>0?fmt(d.reduce((s,r)=>s+r.diferenca,0)):"—"}</span>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })}
              </div>
              <SecLabel>Resultados por Categoria — Time Completo</SecLabel>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16,marginBottom:32}}>
                {(()=>{const d=scopeCat("Reajuste"),rc=revs.filter(r=>r.categoria==="Reajuste");return(
                  <div style={{background:"#13131A",border:`1px solid ${CAT_COM_C["Reajuste"]}44`,padding:20}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
                      <div style={{width:9,height:9,borderRadius:"50%",background:CAT_COM_C["Reajuste"]}}/>
                      <div style={{fontSize:15,color:CAT_COM_C["Reajuste"]}}>Reajuste</div>
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
                      <Mini label="Quantidade" val={d.length}/><Mini label="Revisões" val={rc.length}/>
                      <Mini label="Val. Mensal Atual" val={fmt(d.reduce((s,r)=>s+r.valAtual,0))}/>
                      <Mini label="Val. c/ Reajuste" val={fmt(d.reduce((s,r)=>s+r.valPleito,0))}/>
                    </div>
                    <div style={{background:"#0D0D0F",padding:"8px 12px"}}>
                      <div style={{fontSize:8,letterSpacing:2,color:"#4A4A4A",textTransform:"uppercase",marginBottom:3}}>Aumento do Contrato</div>
                      <div style={{fontSize:16,color:"#C8A96E"}}>{fmt(d.reduce((s,r)=>s+r.diferenca,0))}</div>
                    </div>
                  </div>
                );})()}
                {(()=>{const d=scopeCat("Defesa de Território"),rc=revs.filter(r=>r.categoria==="Defesa de Território");return(
                  <div style={{background:"#13131A",border:`1px solid ${CAT_COM_C["Defesa de Território"]}44`,padding:20}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
                      <div style={{width:9,height:9,borderRadius:"50%",background:CAT_COM_C["Defesa de Território"]}}/>
                      <div style={{fontSize:15,color:CAT_COM_C["Defesa de Território"]}}>Defesa de Território</div>
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                      <Mini label="Quantidade" val={d.length}/><Mini label="Revisões" val={rc.length}/>
                      <Mini label="Val. Mensal" val={fmt(d.reduce((s,r)=>s+r.valAtual,0))}/>
                      <Mini label="Aumento" val={fmt(d.reduce((s,r)=>s+r.diferenca,0))} highlight/>
                    </div>
                  </div>
                );})()}
                {(()=>{
                  const cats=["Up Selling","Alteração de Escopo","Renovação"];
                  const d=scopeAll.filter(r=>cats.includes(r.categoria));
                  const rc=revs.filter(r=>cats.includes(r.categoria));
                  const spot=d.filter(r=>(r.tipoProposta||"").toUpperCase().includes("SPOT")).length;
                  return(
                    <div style={{background:"#13131A",border:`1px solid ${CAT_COM_C["Up Selling"]}44`,padding:20}}>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
                        <div style={{width:9,height:9,borderRadius:"50%",background:CAT_COM_C["Up Selling"]}}/>
                        <div style={{fontSize:15,color:CAT_COM_C["Up Selling"]}}>Aumento / Novos Escopos</div>
                      </div>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
                        <Mini label="Quantidade" val={d.length}/><Mini label="Revisões" val={rc.length}/>
                        <Mini label="Val. Mensal" val={fmt(d.reduce((s,r)=>s+r.valPleito,0))}/>
                        <Mini label="Aumento" val={fmt(d.reduce((s,r)=>s+r.diferenca,0))} highlight/>
                      </div>
                      <div style={{display:"flex",gap:8}}>
                        <div style={{flex:1,background:"#0D0D0F",padding:"8px 12px",textAlign:"center"}}>
                          <div style={{fontSize:8,letterSpacing:2,color:"#4A4A4A",textTransform:"uppercase",marginBottom:3}}>Recorrente</div>
                          <div style={{fontSize:16,color:"#A8E87C"}}>{d.length-spot}</div>
                        </div>
                        <div style={{flex:1,background:"#0D0D0F",padding:"8px 12px",textAlign:"center"}}>
                          <div style={{fontSize:8,letterSpacing:2,color:"#4A4A4A",textTransform:"uppercase",marginBottom:3}}>Spot</div>
                          <div style={{fontSize:16,color:"#C8A96E"}}>{spot}</div>
                        </div>
                      </div>
                    </div>
                  );
                })()}
              </div>
              <SecLabel>Sinalização de Clientes</SecLabel>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16}}>
                {["verde","amarelo","vermelho"].map(sin=>{
                  const list=cStats.filter(c=>c.sin===sin);
                  return(
                    <div key={sin} style={{background:"#13131A",border:`1px solid ${SIN_C[sin]}44`,padding:18}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                        <div style={{fontSize:13,color:SIN_C[sin]}}>{SIN_LB[sin]}</div>
                        <div style={{fontSize:20,color:SIN_C[sin],fontWeight:"bold"}}>{list.length}</div>
                      </div>
                      <div style={{display:"flex",flexDirection:"column",gap:5}}>
                        {list.slice(0,5).map(c=>(
                          <div key={c.grupo} style={{display:"flex",justifyContent:"space-between",padding:"5px 10px",background:SIN_BG[sin]}}>
                            <span style={{fontSize:11,color:"#C8C8C8"}}>{c.grupo}</span>
                            <span style={{fontSize:10,color:SIN_C[sin]}}>{c.rows.length}x</span>
                          </div>
                        ))}
                        {list.length>5&&<div style={{fontSize:10,color:"#4A4A4A",textAlign:"center",paddingTop:4}}>+{list.length-5} clientes</div>}
                        {!list.length&&<div style={{fontSize:11,color:"#3A3A3A",textAlign:"center",padding:"10px 0"}}>Nenhum classificado</div>}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })()}

        {/* ESFORÇO */}
        {fileName&&!loading&&view==="tempo"&&(
          <div>
            <div style={{marginBottom:22}}>
              <SecLabel>Analistas — Esforço</SecLabel>
              <div style={{fontSize:18}}>Onde o Time <span style={{color:"#C8A96E"}}>Gasta Tempo</span></div>
            </div>
            <div style={{background:"#13131A",border:"1px solid #1E1E24",padding:20,marginBottom:18}}>
              <div style={{fontSize:10,letterSpacing:3,color:"#5A5A5A",textTransform:"uppercase",marginBottom:12}}>
                Time — {analistasAtivs.length} atividades · {revs.filter(r=>ANALISTAS.includes(r.responsavel)).length} revisões
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10}}>
                {contarEnt(analistasAtivs).map(({cat,qtd,p})=>(
                  <div key={cat} style={{padding:"10px 14px",background:"#0D0D0F",borderLeft:`3px solid ${CAT_ENT_C[cat]||"#3A3A3A"}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
                      <span style={{fontSize:12,color:CAT_ENT_C[cat]||"#F0EDE8"}}>{cat}</span>
                      <div style={{display:"flex",gap:10}}>
                        <span style={{fontSize:11,color:"#5A5A5A"}}>{qtd}x</span>
                        <span style={{fontSize:14,color:"#F0EDE8",fontWeight:"bold"}}>{p}%</span>
                      </div>
                    </div>
                    <Bar p={p} color={CAT_ENT_C[cat]||"#3A3A3A"}/>
                  </div>
                ))}
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16}}>
              {ANALISTAS.map(p=>{
                const rows=analistasAtivs.filter(r=>r.responsavel===p);
                const cats=contarEnt(rows);
                return(
                  <div key={p} style={{background:"#13131A",border:`1px solid ${COLORS[p]}33`,padding:18}}>
                    <div style={{marginBottom:14}}>
                      <div style={{fontSize:9,letterSpacing:3,color:"#5A5A5A",textTransform:"uppercase",marginBottom:2}}>Analista</div>
                      <div style={{fontSize:15,color:COLORS[p]}}>{p.charAt(0)+p.slice(1).toLowerCase()}</div>
                      <div style={{fontSize:10,color:"#4A4A4A",marginTop:3}}>{rows.length} atividades · {revs.filter(r=>r.responsavel===p).length} revisões</div>
                    </div>
                    {!cats.length&&<div style={{color:"#3A3A3A",fontSize:12}}>Sem dados</div>}
                    {cats.map(({cat,qtd,p:pr})=>(
                      <div key={cat} style={{marginBottom:8}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                          <div style={{display:"flex",alignItems:"center",gap:5}}>
                            <div style={{width:7,height:7,borderRadius:"50%",background:CAT_ENT_C[cat]||"#3A3A3A"}}/>
                            <span style={{fontSize:11,color:"#9A9A9A"}}>{cat}</span>
                          </div>
                          <div style={{display:"flex",gap:7}}>
                            <span style={{fontSize:10,color:"#5A5A5A"}}>{qtd}x</span>
                            <span style={{fontSize:12,color:COLORS[p]}}>{pr}%</span>
                          </div>
                        </div>
                        <Bar p={pr} color={COLORS[p]}/>
                      </div>
                    ))}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* POR CLIENTE */}
        {fileName&&!loading&&view==="clientes"&&(()=>{
          const stats=clienteStats();
          const filtered=fSin==="Todos"?stats:stats.filter(s=>s.sin===fSin);
          return(
            <div>
              <div style={{marginBottom:20}}>
                <SecLabel>Histórico</SecLabel>
                <div style={{fontSize:18}}>Performance <span style={{color:"#C8A96E"}}>por Cliente</span></div>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:18,flexWrap:"wrap"}}>
                <span style={{fontSize:10,letterSpacing:2,color:"#5A5A5A",textTransform:"uppercase"}}>Perfil:</span>
                <Btn label="Todos" active={fSin==="Todos"} onClick={()=>setFSin("Todos")}/>
                {["verde","amarelo","vermelho"].map(s=>(
                  <Btn key={s} label={SIN_LB[s]} active={fSin===s} color={SIN_C[s]} onClick={()=>setFSin(s)}/>
                ))}
                <span style={{marginLeft:"auto",fontSize:11,color:"#5A5A5A"}}>{filtered.length} clientes</span>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:8}}>
                {filtered.map(({grupo,rows,sin,avgReaj,avgPleito,avgAprov,totalAtual,totalPleito,totalDif})=>(
                  <div key={grupo} style={{background:"#13131A",border:"1px solid #1E1E24",padding:"14px 18px",borderLeft:`3px solid ${sin?SIN_C[sin]:"#2A2A2E"}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:10}}>
                      <div>
                        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                          <span style={{fontSize:14}}>{grupo}</span>
                          {sin?<span style={{fontSize:9,padding:"2px 8px",background:SIN_BG[sin],color:SIN_C[sin],border:`1px solid ${SIN_C[sin]}44`}}>{SIN_LB[sin]}</span>
                              :<span style={{fontSize:9,padding:"2px 8px",background:"#1A1A1A",color:"#4A4A4A",border:"1px solid #2A2A2A"}}>⬜ Sem classificação</span>}
                        </div>
                        <div style={{fontSize:10,color:"#5A5A5A"}}>{rows.length} atividades · {rows.filter(r=>r.status.toUpperCase().includes("APROVADO")).length} aprovadas</div>
                      </div>
                      <div style={{display:"flex",gap:14,flexWrap:"wrap"}}>
                        {[
                          {l:"Contrato Atual",v:totalAtual>0?fmt(totalAtual):"—"},
                          {l:"Com Pleito",v:totalPleito>0?fmt(totalPleito):"—"},
                          {l:"Diferença",v:totalDif>0?fmt(totalDif):"—",c:"#C8A96E"},
                          {l:"% Reaj.",v:avgReaj!=null?fmtPct(avgReaj):"—"},
                          {l:"% Pleito",v:avgPleito!=null?fmtPct(avgPleito):"—"},
                          {l:"% Aceito",v:avgAprov!=null?fmtPct(avgAprov):"—",c:avgAprov!=null?"#4A8A4A":undefined},
                        ].map(i=>(
                          <div key={i.l} style={{textAlign:"right"}}>
                            <div style={{fontSize:8,letterSpacing:2,color:"#5A5A5A",textTransform:"uppercase",marginBottom:2}}>{i.l}</div>
                            <div style={{fontSize:13,color:i.c||"#F0EDE8"}}>{i.v}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                ))}
                {!filtered.length&&<div style={{textAlign:"center",padding:"40px 0",color:"#3A3A3A"}}>Nenhum cliente encontrado.</div>}
              </div>
            </div>
          );
        })()}

        {/* HISTÓRICO */}
        {fileName&&!loading&&view==="historico"&&(
          <div>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16,flexWrap:"wrap"}}>
              <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontSize:10,letterSpacing:2,color:"#5A5A5A",textTransform:"uppercase"}}>Resp:</span>
                {["Todos",...ANALISTAS,...LIDERANCA].map(p=>(
                  <Btn key={p} label={p==="Todos"?"Todos":p.charAt(0)+p.slice(1).toLowerCase()} active={fResp===p} color={COLORS[p]||"#C8A96E"} onClick={()=>setFResp(p)}/>
                ))}
              </div>
              <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontSize:10,letterSpacing:2,color:"#5A5A5A",textTransform:"uppercase"}}>Status:</span>
                {["Todos","Aprovado","Em Negociação"].map(s=><Btn key={s} label={s} active={fStat===s} onClick={()=>setFStat(s)}/>)}
              </div>
              <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontSize:10,letterSpacing:2,color:"#5A5A5A",textTransform:"uppercase"}}>Tipo:</span>
                {["Todas",...CAT_COM].map(c=><Btn key={c} label={c} active={fCat===c} color={CAT_COM_C[c]||"#C8A96E"} onClick={()=>setFCat(c)}/>)}
              </div>
              <span style={{marginLeft:"auto",fontSize:11,color:"#5A5A5A"}}>{filtHistorico.length} registros</span>
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                <thead>
                  <tr style={{borderBottom:"1px solid #1E1E24"}}>
                    {["Resp.","Proposta","Cliente","Atividade","Tipo","Val. Atual","Com Pleito","Diferença","% Reaj","% Pleito","% Aceito","Entregáveis","Status","Rev?"].map(h=>(
                      <th key={h} style={{textAlign:"left",padding:"6px 10px",fontSize:8,letterSpacing:2,color:"#4A4A4A",textTransform:"uppercase",whiteSpace:"nowrap"}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtHistorico.map((r,i)=>(
                    <tr key={i} style={{borderBottom:"1px solid #111116",background:i%2===0?"#13131A":"#0F0F16"}}>
                      <td style={{padding:"6px 10px",color:COLORS[r.responsavel]||"#F0EDE8",borderLeft:`2px solid ${COLORS[r.responsavel]||"#3A3A3A"}`,whiteSpace:"nowrap"}}>{r.responsavel.charAt(0)+r.responsavel.slice(1).toLowerCase()}</td>
                      <td style={{padding:"6px 10px",color:"#5A5A5A",whiteSpace:"nowrap"}}>{r.nProposta}</td>
                      <td style={{padding:"6px 10px",maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.cliente}</td>
                      <td style={{padding:"6px 10px",color:"#C8A96E",maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.atividade}</td>
                      <td style={{padding:"6px 10px",whiteSpace:"nowrap"}}><Tag label={r.categoria} color={CAT_COM_C[r.categoria]}/></td>
                      <td style={{padding:"6px 10px",whiteSpace:"nowrap"}}>{r.valAtual>0?fmt(r.valAtual):"—"}</td>
                      <td style={{padding:"6px 10px",whiteSpace:"nowrap"}}>{r.valPleito>0?fmt(r.valPleito):"—"}</td>
                      <td style={{padding:"6px 10px",color:"#C8A96E",whiteSpace:"nowrap"}}>{r.diferenca>0?fmt(r.diferenca):"—"}</td>
                      <td style={{padding:"6px 10px",whiteSpace:"nowrap"}}>{r.pctReaj>0?fmtPct(r.pctReaj):"—"}</td>
                      <td style={{padding:"6px 10px",whiteSpace:"nowrap"}}>{r.pctPleito>0?fmtPct(r.pctPleito):"—"}</td>
                      <td style={{padding:"6px 10px",color:"#4A8A4A",whiteSpace:"nowrap"}}>{r.aprovPct>0?fmtPct(r.aprovPct):"—"}</td>
                      <td style={{padding:"6px 10px"}}>
                        <div style={{display:"flex",flexWrap:"wrap",gap:3}}>
                          {r.entregaveis.map(e=><Tag key={e} label={e} color={CAT_ENT_C[e]||"#5A5A5A"}/>)}
                        </div>
                      </td>
                      <td style={{padding:"6px 10px"}}>
                        <span style={{fontSize:9,padding:"2px 7px",background:r.status.toUpperCase().includes("APROVADO")?"#0A2A0A":"#1A1A0A",color:r.status.toUpperCase().includes("APROVADO")?"#4A8A4A":"#8A8A4A",whiteSpace:"nowrap"}}>{r.status}</span>
                      </td>
                      <td style={{padding:"6px 10px",textAlign:"center",color:r.isRevisao?"#E87C7C":"#3A3A3A"}}>{r.isRevisao?"Rev":"—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!filtHistorico.length&&<div style={{textAlign:"center",padding:"40px 0",color:"#3A3A3A"}}>Nenhum registro encontrado.</div>}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
