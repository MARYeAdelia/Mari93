import { useState, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── EQUIPE ────────────────────────────────────────────────────────────────
const ANALISTAS = ["MARIANA","WILDER","GIOVANNI"];
const LIDERANCA = ["CARLA","DARLAN"];
const EQUIPE    = [...ANALISTAS, ...LIDERANCA];
const COR = {
  MARIANA:"#E8A87C", WILDER:"#7CB9E8", GIOVANNI:"#A8E87C",
  CARLA:"#C8A96E",   DARLAN:"#B87CE8"
};

// ─── CATEGORIAS ────────────────────────────────────────────────────────────
const CATS = ["Reajuste","Upselling","Alteração de Escopo","Defesa de Território","Renovação","BID/Cotação","Outros"];
const COR_CAT = {
  "Reajuste":"#E8A87C","Upselling":"#A8E87C","Alteração de Escopo":"#B87CE8",
  "Defesa de Território":"#C8A96E","Renovação":"#7CB9E8","BID/Cotação":"#E87C9A","Outros":"#5A5A5A"
};

// ─── ENTREGÁVEIS ────────────────────────────────────────────────────────────
const COR_ENT = {
  "PEC":"#7CB9E8","Proposta Comercial":"#C8A96E","Abertura de Custo":"#A8E87C",
  "Carta de Reajuste":"#E8A87C","Notificação de Reajuste":"#B87CE8"
};

// ─── SEMÁFORO ───────────────────────────────────────────────────────────────
const SIN_LABEL = { Verde:"🟢 Boa Negociação", Amarelo:"🟡 Moderada", Vermelho:"🔴 Difícil" };
const SIN_COR   = { Verde:"#3A8A3A", Amarelo:"#8A8A2A", Vermelho:"#8A2A2A" };
const SIN_BG    = { Verde:"#0A2A0A", Amarelo:"#2A2A08", Vermelho:"#2A0A0A" };

// ─── FORMATAÇÃO ─────────────────────────────────────────────────────────────
const brl  = v => (!v && v !== 0) ? "—"
  : new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL",maximumFractionDigits:0}).format(v);
const pct  = v => (v != null && v !== "" && !isNaN(v)) ? `${(parseFloat(v)*100).toFixed(2)}%` : "—";
const num  = v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
const sim  = v => ["SIM","S","X","VERDADEIRO","TRUE"].includes((v||"").toString().toUpperCase().trim());

// ─── HELPERS ────────────────────────────────────────────────────────────────
const detectResp = (strF, strH) => {
  const find = s => EQUIPE.find(p => (s||"").toUpperCase().includes(p));
  return find(strF) || find(strH) || null;
};

const detectCat = tipo => {
  const t = (tipo||"").toUpperCase().trim();
  if (t.includes("REAJUSTE"))                                                    return "Reajuste";
  if (t.includes("UPSELLING")||t.includes("UP-SELLING")||t.includes("UP SELLING")) return "Upselling";
  if (t.includes("DEFESA"))                                                      return "Defesa de Território";
  if (t.includes("ALTERAÇ")||t.includes("REVISÃO DE ESCOPO"))                   return "Alteração de Escopo";
  if (t.includes("RENOVAÇ"))                                                     return "Renovação";
  if (t.includes("BID")||t.includes("COTAÇ")||t.includes("PROJETO"))           return "BID/Cotação";
  return "Outros";
};

const detectSin = v => {
  const s = (v||"").toString();
  if (/verde/i.test(s))   return "Verde";
  if (/amarelo/i.test(s)) return "Amarelo";
  if (/verm/i.test(s))    return "Vermelho";
  return null;
};

// Nº Proposta base: "0023 - 26" ou "0023 - 26 Rev1" → "0023 - 26"
const propBase = np =>
  (np||"").toString().trim().replace(/\s*[-–]?\s*rev\.?\s*\d+$/i,"").trim();

// ─── PARSE EXCEL ────────────────────────────────────────────────────────────
const parseExcel = file => new Promise((res,rej) => {
  const rd = new FileReader();
  rd.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, {type:"array"});
      const nm = wb.SheetNames.find(s => s.toUpperCase() === "DADOS") || wb.SheetNames[0];
      res(XLSX.utils.sheet_to_json(wb.Sheets[nm], {defval:""}));
    } catch(err) { rej(err); }
  };
  rd.onerror = rej;
  rd.readAsArrayBuffer(file);
});

// ─── PROCESSA LINHAS ────────────────────────────────────────────────────────
const processRows = rawRows => {
  if (!rawRows?.length) return [];
  const keys = Object.keys(rawRows[0]);
  const find = (...kws) => keys.find(k =>
    kws.some(kw => k.toUpperCase().replace(/\s+/g," ").trim().includes(kw.toUpperCase()))
  ) || null;
  const get = (r,...kws) => { const k=find(...kws); return k?r[k]:""; };

  return rawRows.map(r => {
    const grupo    = (get(r,"Grupo Cliente")||"").toString().trim();
    const unidade  = (get(r,"Unidade / Filial","Unidade")||"").toString().trim();
    const escopo   = (get(r,"Escopo Atuação","Escopo")||"").toString().trim();
    const respFstr = (get(r,"Responsável Farmer","RESP. FARMER")||"").toString().trim();
    const respHstr = (get(r,"Responsável Hunter","RESP. HUNTER")||"").toString().trim();
    const tipo     = (get(r,"Tipo de Negócio","TIPO DE PROPOSTA")||"").toString().trim();
    const nProp    = (get(r,"Nº Proposta","Nº da Proposta","N PROPOSTA")||"").toString().trim();
    const mes      = (get(r,"Mês","MES")||"").toString().trim();
    const semana   = (get(r,"Semana","SEMANA")||"").toString().trim();
    const obs      = (get(r,"OBS.","OBS")||"").toString().trim();

    const responsavel = detectResp(respFstr, respHstr);
    if (!responsavel || !grupo) return null;

    const eFarmer   = (escopo||"").toUpperCase().includes("FARMER");

    // Nº da Revisão: inteiro (0 = base, ≥1 = revisão)
    const nRevRaw   = get(r,"Nº da Revisão","Nº Revisão");
    const nRev      = parseInt(nRevRaw) || 0;
    const isRevisao = nRev > 0;

    // Chave única por contrato: Grupo + Unidade + Tipo + Nº Proposta base
    // Garante que ADECOAGRO 0023 e 0024 (mesma unidade, mesmo tipo) não colidam
    const contratoKey = `${grupo}||${unidade||grupo}||${tipo}||${propBase(nProp)}`;

    // Financeiro
    const valAtual  = num(get(r,"Valor Contrato Atual"));
    const valPleito = num(get(r,"Valor Final / Pleito"));
    const diferenca = valPleito - valAtual;
    const valProposta = num(get(r,"Valor Proposta","VALOR PROPOSTA"));

    // Percentuais (armazenados em decimal: 0.065 = 6,5%)
    const pctContrato = num(get(r,"% Reajuste Contrato"));
    const pctPleito   = num(get(r,"% Reajuste Pleito"));
    const pctAceito   = num(get(r,"% Reajuste Aceito"));

    const status    = (get(r,"Status","STATUS")||"").toString().trim();
    const semaforo  = detectSin(get(r,"Semáforo (Dificuldade)","Semáforo","Sinalização")||"");
    const dataAprov = (get(r,"Data de Aprovação")||"").toString().trim();

    const fezPec      = get(r,"Fez PEC?","FEZ PEC");
    const fezAbertura = get(r,"Fez Abertura de Custo?","FEZ ABERTURA");
    const fezProposta = get(r,"Fez Proposta Comercial?","FEZ PROPOSTA COMERCIAL");
    const fezCarta    = get(r,"Fez Carta de Reajuste?","FEZ CARTA DE REAJUSTE");
    const fezNotif    = get(r,"Fez Notificação de Reajuste?","FEZ CARTA DE NOTIF","FEZ NOTIF");

    const entregaveis = [
      sim(fezPec)      && "PEC",
      sim(fezAbertura) && "Abertura de Custo",
      sim(fezProposta) && "Proposta Comercial",
      sim(fezCarta)    && "Carta de Reajuste",
      sim(fezNotif)    && "Notificação de Reajuste",
    ].filter(Boolean);

    return {
      grupo, unidade: unidade||grupo, escopo, eFarmer,
      responsavel, respFarmer: respFstr, respHunter: respHstr,
      nProp, tipo, categoria: detectCat(tipo),
      mes, semana, obs, dataAprov,
      nRev, isRevisao, contratoKey,
      valAtual, valPleito, diferenca, valProposta,
      pctContrato, pctPleito, pctAceito,
      status, semaforo, entregaveis,
    };
  }).filter(Boolean);
};

// ─── MOTOR FINANCEIRO ───────────────────────────────────────────────────────
// Por contratoKey → linha com maior nRev
const ultimaPorContrato = rows => {
  const by = {};
  rows.forEach(r => {
    if (!by[r.contratoKey] || r.nRev > by[r.contratoKey].nRev)
      by[r.contratoKey] = r;
  });
  return Object.values(by);
};

// ─── COMPONENTES ────────────────────────────────────────────────────────────
const Btn = ({label,active,color="#C8A96E",onClick}) => (
  <button onClick={onClick} style={{
    padding:"4px 12px",border:"1px solid",cursor:"pointer",fontFamily:"inherit",
    borderColor:active?color:"#2A2A2E",background:active?`${color}18`:"transparent",
    color:active?color:"#5A5A5A",fontSize:10,letterSpacing:2,textTransform:"uppercase",transition:"all .15s"
  }}>{label}</button>
);
const KpiCard = ({label,valor,destaque}) => (
  <div style={{background:"#0D0D0F",padding:"8px 12px"}}>
    <div style={{fontSize:8,letterSpacing:2,color:"#4A4A4A",textTransform:"uppercase",marginBottom:3}}>{label}</div>
    <div style={{fontSize:11,color:destaque?"#C8A96E":"#F0EDE8"}}>{valor||"—"}</div>
  </div>
);
const Tag = ({label,color}) => (
  <span style={{fontSize:9,padding:"2px 7px",background:`${color}22`,color,
                border:`1px solid ${color}44`,whiteSpace:"nowrap",borderRadius:2}}>{label}</span>
);
const Barra = ({p,color}) => (
  <div style={{flex:1,height:5,background:"#1A1A22",borderRadius:3,overflow:"hidden"}}>
    <div style={{width:`${Math.min(p,100)}%`,height:"100%",background:color,transition:"width .4s"}}/>
  </div>
);
const Sec = ({children}) => (
  <div style={{fontSize:9,letterSpacing:4,color:"#5A5A5A",textTransform:"uppercase",marginBottom:10}}>{children}</div>
);

// ─── APP ────────────────────────────────────────────────────────────────────
export default function FarmDashboard() {
  const [data,    setData]    = useState([]);
  const [view,    setView]    = useState("visao");
  const [fResp,   setFResp]   = useState("Todos");
  const [fStat,   setFStat]   = useState("Todos");
  const [fCat,    setFCat]    = useState("Todas");
  const [fSin,    setFSin]    = useState("Todos");
  const [fMes,    setFMes]    = useState("Todos");
  const [fSemana, setFSemana] = useState("Todas");
  const [arquivo, setArquivo] = useState(null);
  const [loading, setLoading] = useState(false);
  const [drag,    setDrag]    = useState(false);
  const [erro,    setErro]    = useState(null);

  const carregar = useCallback(async file => {
    setLoading(true); setErro(null);
    try { setData(processRows(await parseExcel(file))); setArquivo(file.name); }
    catch { setErro("Erro ao ler o arquivo. Verifique se é um .xlsx válido."); }
    finally { setLoading(false); }
  }, []);

  // ── Filtros ──────────────────────────────────────────────────────────────
  const filtraMesSemana = rows => {
    let r = rows;
    if (fMes    !== "Todos")  r = r.filter(x => x.mes    === fMes);
    if (fSemana !== "Todas")  r = r.filter(x => x.semana === fSemana);
    return r;
  };
  const filtraStatus = rows => {
    if (fStat === "Aprovado")      return rows.filter(r => r.status.toUpperCase().includes("APROVADO"));
    if (fStat === "Em Negociação") return rows.filter(r => r.status.toUpperCase().includes("NEGOCI"));
    if (fStat === "Recusado")      return rows.filter(r => r.status.toUpperCase().includes("RECUS"));
    return rows;
  };
  const filtraGeral = rows => {
    let r = rows;
    if (fResp !== "Todos") r = r.filter(x => x.responsavel === fResp);
    r = filtraStatus(r);
    if (fCat !== "Todas")  r = r.filter(x => x.categoria === fCat);
    return r;
  };

  // ── Listas de filtro ─────────────────────────────────────────────────────
  const ORDEM_MES = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"];
  const getMeses = () => {
    const s = new Set(data.map(r=>r.mes).filter(Boolean));
    return ["Todos",...[...s].sort((a,b) =>
      ORDEM_MES.indexOf(a.toUpperCase().substring(0,3)) - ORDEM_MES.indexOf(b.toUpperCase().substring(0,3))
    )];
  };
  const getSemanas = () => {
    const base = fMes==="Todos" ? data : data.filter(r=>r.mes===fMes);
    const s = new Set(base.map(r=>r.semana).filter(Boolean));
    const ord = ["SEMANA 0","SEMANA 1","SEMANA 2","SEMANA 3","SEMANA 4","SEMANA 5"];
    return ["Todas",...[...s].sort((a,b)=>ord.indexOf(a.toUpperCase())-ord.indexOf(b.toUpperCase()))];
  };
  const getResps = () => {
    const s = new Set(data.map(r=>r.responsavel).filter(Boolean));
    return ["Todos",...EQUIPE.filter(p=>s.has(p))];
  };

  // ── Bases ────────────────────────────────────────────────────────────────
  const periodoData = filtraMesSemana(data);
  const linhasFilt  = filtraMesSemana(filtraGeral(data));

  // Motor financeiro: última revisão por contrato, só farmer, com valor
  const ultimasFarmer = ultimaPorContrato(
    filtraStatus(periodoData).filter(r => r.eFarmer && r.valAtual > 0)
  );

  // Stats por grupo cliente
  const statsGrupos = () => {
    const by = {};
    periodoData.forEach(r => {
      if (!by[r.grupo]) by[r.grupo] = { rows:[], sin:null };
      by[r.grupo].rows.push(r);
      if (!by[r.grupo].sin && r.semaforo) by[r.grupo].sin = r.semaforo;
    });
    return Object.entries(by).map(([grupo,{rows,sin}]) => {
      const ults  = ultimaPorContrato(filtraStatus(rows).filter(r=>r.eFarmer&&r.valAtual>0));
      const somaA = ults.reduce((s,r)=>s+r.valAtual,0);
      const somaP = ults.reduce((s,r)=>s+r.valPleito,0);
      const taxa  = somaA>0 ? (somaP-somaA)/somaA : null;
      return { grupo, rows, sin, ults, somaAtual:somaA, somaPleito:somaP, somaDif:somaP-somaA, taxa };
    }).sort((a,b)=>b.rows.length-a.rows.length);
  };

  const contarEnt = rows => {
    const c={}; let t=0;
    rows.forEach(r=>r.entregaveis.forEach(e=>{c[e]=(c[e]||0)+1;t++;}));
    return Object.entries(c).map(([cat,qtd])=>({cat,qtd,p:t?Math.round(qtd/t*100):0})).sort((a,b)=>b.qtd-a.qtd);
  };

  // ─── RENDER ──────────────────────────────────────────────────────────────
  return (
    <div style={{minHeight:"100vh",background:"#0D0D0F",color:"#F0EDE8",fontFamily:"'Georgia',serif"}}>

      {/* HEADER */}
      <div style={{borderBottom:"1px solid #1E1E24",padding:"16px 28px",display:"flex",
                   alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
        <div>
          <div style={{fontSize:10,letterSpacing:4,color:"#5A5A5A",textTransform:"uppercase",marginBottom:2}}>Time Comercial</div>
          <div style={{fontSize:20}}>Farm <span style={{color:"#C8A96E"}}>Performance</span></div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          {arquivo&&<div style={{fontSize:10,color:"#4A4A4A",background:"#13131A",border:"1px solid #1E1E24",padding:"4px 10px"}}>📄 {arquivo}</div>}
          <label style={{cursor:"pointer"}}>
            <input type="file" accept=".xlsx,.xls" onChange={e=>{const f=e.target.files[0];if(f)carregar(f);}} style={{display:"none"}}/>
            <div style={{padding:"6px 16px",border:"1px solid #C8A96E",color:"#C8A96E",fontSize:10,letterSpacing:2,textTransform:"uppercase",cursor:"pointer"}}>
              {arquivo?"Trocar Planilha":"Carregar Planilha"}
            </div>
          </label>
          {[["visao","Visão Geral"],["esforco","Esforço"],["clientes","Por Cliente"],["historico","Histórico"]].map(([v,l])=>(
            <Btn key={v} label={l} active={view===v} onClick={()=>setView(v)}/>
          ))}
        </div>
      </div>

      <div style={{padding:"22px 28px"}}>

        {/* DROP ZONE */}
        {!arquivo&&(
          <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
               onDrop={e=>{e.preventDefault();setDrag(false);const f=e.dataTransfer.files[0];if(f)carregar(f);}}
               style={{border:`2px dashed ${drag?"#C8A96E":"#2A2A2E"}`,background:drag?"#C8A96E08":"#13131A",
                       padding:"70px 40px",textAlign:"center",cursor:"pointer"}}>
            <label style={{cursor:"pointer"}}>
              <input type="file" accept=".xlsx,.xls" onChange={e=>{const f=e.target.files[0];if(f)carregar(f);}} style={{display:"none"}}/>
              <div style={{fontSize:34,marginBottom:12}}>📊</div>
              <div style={{fontSize:14,color:"#C8A96E",marginBottom:6}}>Arraste a planilha aqui</div>
              <div style={{fontSize:11,color:"#4A4A4A"}}>Controle_Comercial_Farmer.xlsx — aba "Dados" lida automaticamente</div>
            </label>
          </div>
        )}
        {loading&&<div style={{textAlign:"center",padding:"60px 0",color:"#C8A96E"}}>Lendo planilha...</div>}
        {erro&&<div style={{background:"#2A1010",border:"1px solid #5A2020",padding:"12px 18px",color:"#E87C7C",fontSize:13}}>{erro}</div>}

        {/* FILTROS GLOBAIS */}
        {arquivo&&!loading&&(
          <div style={{display:"flex",gap:14,marginBottom:20,flexWrap:"wrap",alignItems:"center"}}>
            <div style={{display:"flex",gap:5,alignItems:"center"}}>
              <span style={{fontSize:9,letterSpacing:2,color:"#5A5A5A",textTransform:"uppercase"}}>Status:</span>
              {["Todos","Aprovado","Em Negociação","Recusado"].map(s=><Btn key={s} label={s} active={fStat===s} onClick={()=>setFStat(s)}/>)}
            </div>
            <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:9,letterSpacing:2,color:"#5A5A5A",textTransform:"uppercase"}}>Mês:</span>
              {getMeses().map(m=><Btn key={m} label={m} active={fMes===m} color="#7CB9E8" onClick={()=>{setFMes(m);setFSemana("Todas");}}/>)}
            </div>
            <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:9,letterSpacing:2,color:"#5A5A5A",textTransform:"uppercase"}}>Semana:</span>
              {getSemanas().map(s=><Btn key={s} label={s} active={fSemana===s} color="#A8E87C" onClick={()=>setFSemana(s)}/>)}
            </div>
          </div>
        )}

        {/* ══ VISÃO GERAL ══════════════════════════════════════════════ */}
        {arquivo&&!loading&&view==="visao"&&(()=>{
          const somaAtual  = ultimasFarmer.reduce((s,r)=>s+r.valAtual,0);
          const somaPleito = ultimasFarmer.reduce((s,r)=>s+r.valPleito,0);
          const somaDif    = somaPleito - somaAtual;

          return (
            <div>
              <Sec>Totais do Time</Sec>
              <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12,marginBottom:28}}>
                {[
                  {l:"Atividades",        v: periodoData.filter(r=>!r.isRevisao).length},
                  {l:"Revisões",          v: periodoData.filter(r=>r.isRevisao).length},
                  {l:"Valor Atual",       v: brl(somaAtual)},
                  {l:"Valor c/ Pleito",   v: brl(somaPleito)},
                  {l:"Diferença (Ganho)", v: brl(somaDif)},
                ].map(i=>(
                  <div key={i.l} style={{background:"#13131A",border:"1px solid #1E1E24",padding:"14px 18px"}}>
                    <div style={{fontSize:9,letterSpacing:3,color:"#5A5A5A",textTransform:"uppercase",marginBottom:4}}>{i.l}</div>
                    <div style={{fontSize:20,color:"#C8A96E"}}>{i.v}</div>
                  </div>
                ))}
              </div>

              <Sec>Analistas</Sec>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16,marginBottom:32}}>
                {ANALISTAS.filter(p=>periodoData.some(r=>r.responsavel===p)).map(p=>{
                  const ultP  = ultimaPorContrato(filtraStatus(periodoData).filter(r=>r.responsavel===p&&r.eFarmer&&r.valAtual>0));
                  const somaA = ultP.reduce((s,r)=>s+r.valAtual,0);
                  const somaP = ultP.reduce((s,r)=>s+r.valPleito,0);
                  const aprov = ultP.filter(r=>r.status.toUpperCase().includes("APROVADO")).length;
                  const todasP= periodoData.filter(r=>r.responsavel===p);
                  const hunterP=todasP.filter(r=>!r.eFarmer);
                  const porCat= CATS.map(cat=>({cat,ults:ultP.filter(r=>r.categoria===cat)})).filter(x=>x.ults.length);

                  return (
                    <div key={p} style={{background:"#13131A",border:`1px solid ${COR[p]}33`,padding:20}}>
                      <div style={{display:"flex",justifyContent:"space-between",marginBottom:14}}>
                        <div>
                          <div style={{fontSize:9,letterSpacing:3,color:"#5A5A5A",textTransform:"uppercase",marginBottom:2}}>Analista</div>
                          <div style={{fontSize:17,color:COR[p]}}>{p[0]+p.slice(1).toLowerCase()}</div>
                        </div>
                        <div style={{textAlign:"right",fontSize:10}}>
                          <div style={{color:"#4A8A4A"}}>✓ {aprov} aprovados</div>
                          <div style={{color:"#5A5A3A",marginTop:2}}>{todasP.filter(r=>r.isRevisao).length} revisões</div>
                        </div>
                      </div>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:14}}>
                        <KpiCard label="Contrato Atual" valor={brl(somaA)}/>
                        <KpiCard label="Com Pleito"     valor={brl(somaP)}/>
                        <KpiCard label="Diferença"      valor={brl(somaP-somaA)} destaque/>
                      </div>
                      {porCat.length>0&&(
                        <div style={{borderTop:"1px solid #1A1A20",paddingTop:12}}>
                          <div style={{fontSize:9,letterSpacing:2,color:"#4A4A4A",textTransform:"uppercase",marginBottom:8}}>Por Tipo</div>
                          {porCat.map(({cat,ults:d})=>(
                            <div key={cat} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                              <div style={{display:"flex",alignItems:"center",gap:6}}>
                                <div style={{width:7,height:7,borderRadius:"50%",background:COR_CAT[cat]}}/>
                                <span style={{fontSize:11,color:"#8A8A8A"}}>{cat}</span>
                              </div>
                              <div style={{display:"flex",gap:10}}>
                                <span style={{fontSize:11,color:COR[p]}}>{d.length}x</span>
                                <span style={{fontSize:11,color:"#C8A96E",minWidth:65,textAlign:"right"}}>
                                  {d.reduce((s,r)=>s+r.diferenca,0)>0?brl(d.reduce((s,r)=>s+r.diferenca,0)):"—"}
                                </span>
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                      {hunterP.length>0&&(
                        <div style={{borderTop:"1px solid #1A1A20",paddingTop:10,marginTop:10}}>
                          <div style={{fontSize:9,letterSpacing:2,color:"#7CB9E8",textTransform:"uppercase",marginBottom:4}}>🤝 Apoio Hunter</div>
                          <div style={{fontSize:11,color:"#5A5A5A"}}>{hunterP.length} interação{hunterP.length!==1?"ões":""}</div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {LIDERANCA.some(p=>periodoData.some(r=>r.responsavel===p))&&(
                <>
                  <Sec>Liderança</Sec>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:16,marginBottom:32}}>
                    {LIDERANCA.filter(p=>periodoData.some(r=>r.responsavel===p)).map(p=>{
                      const ultP  = ultimaPorContrato(filtraStatus(periodoData).filter(r=>r.responsavel===p&&r.eFarmer&&r.valAtual>0));
                      const somaA = ultP.reduce((s,r)=>s+r.valAtual,0);
                      const somaP = ultP.reduce((s,r)=>s+r.valPleito,0);
                      const aprov = ultP.filter(r=>r.status.toUpperCase().includes("APROVADO")).length;
                      const todasP= periodoData.filter(r=>r.responsavel===p);
                      return (
                        <div key={p} style={{background:"#13131A",border:`1px solid ${COR[p]}33`,padding:20}}>
                          <div style={{display:"flex",justifyContent:"space-between",marginBottom:14}}>
                            <div>
                              <div style={{fontSize:9,letterSpacing:3,color:"#5A5A5A",textTransform:"uppercase",marginBottom:2}}>Liderança</div>
                              <div style={{fontSize:17,color:COR[p]}}>{p[0]+p.slice(1).toLowerCase()}</div>
                            </div>
                            <div style={{textAlign:"right",fontSize:10}}>
                              <div style={{color:"#4A8A4A"}}>✓ {aprov} aprovados</div>
                              <div style={{color:"#5A5A3A",marginTop:2}}>{todasP.filter(r=>r.isRevisao).length} revisões</div>
                            </div>
                          </div>
                          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6,marginBottom:14}}>
                            <KpiCard label="Interações" valor={todasP.length}/>
                            <KpiCard label="Atual"      valor={brl(somaA)}/>
                            <KpiCard label="Com Pleito" valor={brl(somaP)}/>
                            <KpiCard label="Diferença"  valor={brl(somaP-somaA)} destaque/>
                          </div>
                          <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                            {CATS.map(cat=>{const d=ultP.filter(r=>r.categoria===cat);return d.length?<Tag key={cat} label={`${cat} ${d.length}x`} color={COR_CAT[cat]}/>:null;})}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </>
              )}

              <Sec>Semáforo de Clientes</Sec>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16}}>
                {["Verde","Amarelo","Vermelho"].map(sin=>{
                  const grupos=statsGrupos().filter(g=>g.sin===sin);
                  return (
                    <div key={sin} style={{background:"#13131A",border:`1px solid ${SIN_COR[sin]}44`,padding:18}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                        <div style={{fontSize:13,color:SIN_COR[sin]}}>{SIN_LABEL[sin]}</div>
                        <div style={{fontSize:20,color:SIN_COR[sin],fontWeight:"bold"}}>{grupos.length}</div>
                      </div>
                      <div style={{display:"flex",flexDirection:"column",gap:5}}>
                        {grupos.slice(0,6).map(g=>(
                          <div key={g.grupo} style={{display:"flex",justifyContent:"space-between",padding:"5px 10px",background:SIN_BG[sin]}}>
                            <span style={{fontSize:11,color:"#C8C8C8"}}>{g.grupo}</span>
                            <span style={{fontSize:10,color:SIN_COR[sin]}}>{g.ults.length} contrato{g.ults.length!==1?"s":""}</span>
                          </div>
                        ))}
                        {grupos.length>6&&<div style={{fontSize:10,color:"#4A4A4A",textAlign:"center",paddingTop:4}}>+{grupos.length-6} grupos</div>}
                        {!grupos.length&&<div style={{fontSize:11,color:"#3A3A3A",textAlign:"center",padding:"10px 0"}}>Nenhum classificado</div>}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })()}

        {/* ══ ESFORÇO ══════════════════════════════════════════════════ */}
        {arquivo&&!loading&&view==="esforco"&&(()=>{
          const analistas=periodoData.filter(r=>ANALISTAS.includes(r.responsavel));
          return (
            <div>
              <div style={{marginBottom:22}}>
                <Sec>Esforço Operacional</Sec>
                <div style={{fontSize:18}}>Mesa de <span style={{color:"#C8A96E"}}>Trabalho</span></div>
                <div style={{fontSize:11,color:"#4A4A4A",marginTop:4}}>{analistas.length} interações no período</div>
              </div>
              <div style={{background:"#13131A",border:"1px solid #1E1E24",padding:20,marginBottom:18}}>
                <div style={{fontSize:10,letterSpacing:3,color:"#5A5A5A",textTransform:"uppercase",marginBottom:12}}>Distribuição do Time</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10}}>
                  {contarEnt(analistas).map(({cat,qtd,p})=>(
                    <div key={cat} style={{padding:"10px 14px",background:"#0D0D0F",borderLeft:`3px solid ${COR_ENT[cat]||"#3A3A3A"}`}}>
                      <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
                        <span style={{fontSize:12,color:COR_ENT[cat]||"#F0EDE8"}}>{cat}</span>
                        <div style={{display:"flex",gap:10}}>
                          <span style={{fontSize:11,color:"#5A5A5A"}}>{qtd}x</span>
                          <span style={{fontSize:14,color:"#F0EDE8",fontWeight:"bold"}}>{p}%</span>
                        </div>
                      </div>
                      <Barra p={p} color={COR_ENT[cat]||"#3A3A3A"}/>
                    </div>
                  ))}
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16}}>
                {ANALISTAS.filter(p=>periodoData.some(r=>r.responsavel===p)).map(p=>{
                  const rows=periodoData.filter(r=>r.responsavel===p);
                  const cats=contarEnt(rows);
                  return (
                    <div key={p} style={{background:"#13131A",border:`1px solid ${COR[p]}33`,padding:18}}>
                      <div style={{marginBottom:14}}>
                        <div style={{fontSize:9,letterSpacing:3,color:"#5A5A5A",textTransform:"uppercase",marginBottom:2}}>Analista</div>
                        <div style={{fontSize:15,color:COR[p]}}>{p[0]+p.slice(1).toLowerCase()}</div>
                        <div style={{fontSize:10,color:"#4A4A4A",marginTop:3}}>{rows.length} interações · {rows.filter(r=>r.isRevisao).length} revisões</div>
                      </div>
                      {!cats.length&&<div style={{color:"#3A3A3A",fontSize:12}}>Sem checklist preenchido</div>}
                      {cats.map(({cat,qtd,p:pr})=>(
                        <div key={cat} style={{marginBottom:8}}>
                          <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                            <div style={{display:"flex",alignItems:"center",gap:5}}>
                              <div style={{width:7,height:7,borderRadius:"50%",background:COR_ENT[cat]||"#3A3A3A"}}/>
                              <span style={{fontSize:11,color:"#9A9A9A"}}>{cat}</span>
                            </div>
                            <div style={{display:"flex",gap:7}}>
                              <span style={{fontSize:10,color:"#5A5A5A"}}>{qtd}x</span>
                              <span style={{fontSize:12,color:COR[p]}}>{pr}%</span>
                            </div>
                          </div>
                          <Barra p={pr} color={COR[p]}/>
                        </div>
                      ))}
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })()}

        {/* ══ POR CLIENTE ══════════════════════════════════════════════ */}
        {arquivo&&!loading&&view==="clientes"&&(()=>{
          const stats   =statsGrupos();
          const filtered=fSin==="Todos"?stats:stats.filter(g=>g.sin===fSin);
          const TH=({children,right})=>(
            <th style={{textAlign:right?"right":"left",padding:"5px 10px",fontSize:8,letterSpacing:2,
                        color:"#4A4A4A",textTransform:"uppercase",whiteSpace:"nowrap",fontWeight:"normal",
                        borderBottom:"1px solid #111116"}}>{children}</th>
          );
          const TD=({children,color,right})=>(
            <td style={{padding:"6px 10px",fontSize:11,textAlign:right?"right":"left",
                        color:color||"#C8C8C8",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{children||"—"}</td>
          );
          return (
            <div>
              <div style={{marginBottom:20}}>
                <Sec>Por Cliente</Sec>
                <div style={{fontSize:18}}>Performance <span style={{color:"#C8A96E"}}>por Grupo</span></div>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:18,flexWrap:"wrap"}}>
                <span style={{fontSize:9,letterSpacing:2,color:"#5A5A5A",textTransform:"uppercase"}}>Semáforo:</span>
                <Btn label="Todos" active={fSin==="Todos"} onClick={()=>setFSin("Todos")}/>
                {["Verde","Amarelo","Vermelho"].map(s=>(
                  <Btn key={s} label={SIN_LABEL[s]} active={fSin===s} color={SIN_COR[s]} onClick={()=>setFSin(s)}/>
                ))}
                <span style={{marginLeft:"auto",fontSize:11,color:"#5A5A5A"}}>{filtered.length} grupos</span>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:12}}>
                {filtered.map(({grupo,rows,sin,ults,somaAtual,somaPleito,somaDif,taxa})=>(
                  <div key={grupo} style={{background:"#13131A",border:"1px solid #1E1E24",
                                           borderLeft:`3px solid ${sin?SIN_COR[sin]:"#2A2A2E"}`}}>
                    {/* Cabeçalho */}
                    <div style={{padding:"12px 18px",display:"flex",alignItems:"center",
                                 justifyContent:"space-between",flexWrap:"wrap",gap:10,
                                 borderBottom:ults.length?"1px solid #1A1A20":"none"}}>
                      <div style={{display:"flex",alignItems:"center",gap:10}}>
                        <span style={{fontSize:15}}>{grupo}</span>
                        {sin
                          ?<span style={{fontSize:9,padding:"2px 8px",background:SIN_BG[sin],color:SIN_COR[sin],border:`1px solid ${SIN_COR[sin]}44`}}>{SIN_LABEL[sin]}</span>
                          :<span style={{fontSize:9,padding:"2px 8px",background:"#1A1A1A",color:"#4A4A4A",border:"1px solid #2A2A2A"}}>⬜ Sem semáforo</span>}
                        <span style={{fontSize:10,color:"#4A4A4A"}}>{ults.length} unidade{ults.length!==1?"s":""}</span>
                      </div>
                      {somaAtual>0&&(
                        <div style={{display:"flex",gap:20}}>
                          {[
                            {l:"Total Atual",  v:brl(somaAtual)},
                            {l:"Total Pleito", v:brl(somaPleito)},
                            {l:"Ganho",        v:brl(somaDif),              c:"#C8A96E"},
                            {l:"% Ponderado",  v:taxa!=null?pct(taxa):"—",  c:"#A8E87C"},
                          ].map(i=>(
                            <div key={i.l} style={{textAlign:"right"}}>
                              <div style={{fontSize:8,letterSpacing:2,color:"#4A4A4A",textTransform:"uppercase",marginBottom:2}}>{i.l}</div>
                              <div style={{fontSize:13,color:i.c||"#F0EDE8"}}>{i.v}</div>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                    {/* Tabela de unidades */}
                    {ults.length>0&&(
                      <table style={{width:"100%",borderCollapse:"collapse"}}>
                        <thead>
                          <tr style={{background:"#0D0D0F"}}>
                            <TH>Unidade / Filial</TH>
                            <TH>Tipo</TH>
                            <TH>Rev.</TH>
                            <TH right>Contrato Atual</TH>
                            <TH right>Com Pleito</TH>
                            <TH right>Ganho</TH>
                            <TH right>% Reajuste</TH>
                            <TH right>% Aceito</TH>
                            <TH>Status</TH>
                          </tr>
                        </thead>
                        <tbody>
                          {ults.sort((a,b)=>a.unidade.localeCompare(b.unidade)).map((u,i)=>{
                            const ganho  =u.valPleito-u.valAtual;
                            const pctReaj=u.valAtual>0?ganho/u.valAtual:null;
                            const isAprov=u.status.toUpperCase().includes("APROVADO");
                            const isRecus=u.status.toUpperCase().includes("RECUS");
                            return (
                              <tr key={i} style={{borderBottom:"1px solid #111116",background:i%2===0?"#13131A":"#111118"}}>
                                <TD color="#F0EDE8">{u.unidade}</TD>
                                <td style={{padding:"6px 10px"}}><Tag label={u.categoria} color={COR_CAT[u.categoria]||"#5A5A5A"}/></td>
                                <TD color="#5A5A5A" right>{u.nRev}</TD>
                                <TD right color="#8A8A8A">{u.valAtual>0?brl(u.valAtual):"—"}</TD>
                                <TD right>{u.valPleito>0?brl(u.valPleito):"—"}</TD>
                                <TD right color={ganho>0?"#C8A96E":"#5A5A5A"}>{ganho>0?brl(ganho):"—"}</TD>
                                <TD right color="#7A7A9A">{pctReaj!=null?pct(pctReaj):"—"}</TD>
                                <TD right color="#A8E87C">{u.pctAceito>0?pct(u.pctAceito):"—"}</TD>
                                <td style={{padding:"6px 10px"}}>
                                  <span style={{fontSize:9,padding:"2px 6px",whiteSpace:"nowrap",
                                    background:isAprov?"#0A2A0A":isRecus?"#2A0A0A":"#1A1A0A",
                                    color:isAprov?"#4A8A4A":isRecus?"#8A4A4A":"#8A8A4A"}}>
                                    {u.status||"—"}
                                  </span>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    )}
                  </div>
                ))}
                {!filtered.length&&<div style={{textAlign:"center",padding:"40px 0",color:"#3A3A3A"}}>Nenhum grupo encontrado.</div>}
              </div>
            </div>
          );
        })()}

        {/* ══ HISTÓRICO ════════════════════════════════════════════════ */}
        {arquivo&&!loading&&view==="historico"&&(
          <div>
            <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
              <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontSize:9,letterSpacing:2,color:"#5A5A5A",textTransform:"uppercase"}}>Resp:</span>
                {getResps().map(p=>(
                  <Btn key={p} label={p==="Todos"?"Todos":p[0]+p.slice(1).toLowerCase()}
                       active={fResp===p} color={COR[p]||"#C8A96E"} onClick={()=>setFResp(p)}/>
                ))}
              </div>
              <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontSize:9,letterSpacing:2,color:"#5A5A5A",textTransform:"uppercase"}}>Tipo:</span>
                {["Todas",...CATS].map(c=><Btn key={c} label={c} active={fCat===c} color={COR_CAT[c]||"#C8A96E"} onClick={()=>setFCat(c)}/>)}
              </div>
              <span style={{marginLeft:"auto",fontSize:11,color:"#5A5A5A"}}>{linhasFilt.length} registros</span>
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                <thead>
                  <tr style={{borderBottom:"1px solid #1E1E24"}}>
                    {["Mês","Sem.","Resp.","Escopo","Grupo","Unidade","Tipo","Rev.","Contrato Atual","Com Pleito","Ganho","% Reajuste","% Aceito","Semáforo","Status","Obs"].map(h=>(
                      <th key={h} style={{textAlign:"left",padding:"6px 10px",fontSize:8,letterSpacing:2,
                                          color:"#4A4A4A",textTransform:"uppercase",whiteSpace:"nowrap",fontWeight:"normal"}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {linhasFilt.map((r,i)=>{
                    const ganho  =r.valPleito-r.valAtual;
                    const pctReaj=r.valAtual>0?ganho/r.valAtual:null;
                    const isAprov=r.status.toUpperCase().includes("APROVADO");
                    const isRecus=r.status.toUpperCase().includes("RECUS");
                    return (
                      <tr key={i} style={{borderBottom:"1px solid #111116",background:i%2===0?"#13131A":"#0F0F16"}}>
                        <td style={{padding:"6px 10px",color:"#5A5A5A",whiteSpace:"nowrap"}}>{r.mes}</td>
                        <td style={{padding:"6px 10px",color:"#5A5A5A",whiteSpace:"nowrap",fontSize:10}}>{r.semana}</td>
                        <td style={{padding:"6px 10px",color:COR[r.responsavel]||"#F0EDE8",
                                    borderLeft:`2px solid ${COR[r.responsavel]||"#3A3A3A"}`,whiteSpace:"nowrap"}}>
                          {r.responsavel[0]+r.responsavel.slice(1).toLowerCase()}
                        </td>
                        <td style={{padding:"6px 10px"}}>
                          <span style={{fontSize:9,padding:"2px 5px",
                            background:r.eFarmer?"#0A2A0A":"#0A1A2A",color:r.eFarmer?"#4A8A4A":"#3A6A8A"}}>
                            {r.eFarmer?"Farmer":"Hunter"}
                          </span>
                        </td>
                        <td style={{padding:"6px 10px",maxWidth:100,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#8A8A8A"}}>{r.grupo}</td>
                        <td style={{padding:"6px 10px",maxWidth:130,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#C8A96E"}}>{r.unidade}</td>
                        <td style={{padding:"6px 10px",whiteSpace:"nowrap"}}><Tag label={r.categoria} color={COR_CAT[r.categoria]||"#5A5A5A"}/></td>
                        <td style={{padding:"6px 10px",textAlign:"center",color:r.isRevisao?"#E87C7C":"#3A3A3A",fontSize:10}}>{r.nRev}</td>
                        <td style={{padding:"6px 10px",color:"#8A8A8A",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{r.valAtual>0?brl(r.valAtual):"—"}</td>
                        <td style={{padding:"6px 10px",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{r.valPleito>0?brl(r.valPleito):"—"}</td>
                        <td style={{padding:"6px 10px",color:ganho>0?"#C8A96E":"#5A5A5A",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{ganho>0?brl(ganho):"—"}</td>
                        <td style={{padding:"6px 10px",color:"#7A7A9A",whiteSpace:"nowrap"}}>{pctReaj!=null?pct(pctReaj):"—"}</td>
                        <td style={{padding:"6px 10px",color:"#A8E87C",whiteSpace:"nowrap"}}>{r.pctAceito>0?pct(r.pctAceito):"—"}</td>
                        <td style={{padding:"6px 10px"}}>
                          {r.semaforo&&<span style={{fontSize:9,padding:"2px 5px",background:SIN_BG[r.semaforo],color:SIN_COR[r.semaforo]}}>{r.semaforo}</span>}
                        </td>
                        <td style={{padding:"6px 10px"}}>
                          <span style={{fontSize:9,padding:"2px 6px",whiteSpace:"nowrap",
                            background:isAprov?"#0A2A0A":isRecus?"#2A0A0A":"#1A1A0A",
                            color:isAprov?"#4A8A4A":isRecus?"#8A4A4A":"#8A8A4A"}}>
                            {r.status||"—"}
                          </span>
                        </td>
                        <td style={{padding:"6px 10px",maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#4A4A4A",fontSize:10}}>{r.obs}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {!linhasFilt.length&&<div style={{textAlign:"center",padding:"40px 0",color:"#3A3A3A"}}>Nenhum registro encontrado.</div>}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
