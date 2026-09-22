/* Dashboard de Oportunidades — GitHub Pages + Supabase */
const CFG = window.APP_CONFIG || {};
const db = window.supabase.createClient(CFG.SUPABASE_URL, CFG.SUPABASE_KEY);

const REQUIRED = ["Código","Nome da conta","Tipo de oportunidade","Etapa","Probabilidade","Data prevista (dias)","Responsável","Valor SAAS","Valor CDU/Adesão","Valor SMS","Valor Serviços Não Recorrentes"];
const MES = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
const TODAY = new Date();
const CUR = TODAY.getFullYear()*12 + TODAY.getMonth();
const $ = id => document.getElementById(id);

let rows = [], ultimaCarga = null, pendente = null, papel = "usuario";
let state = { hz:"pp", exec:"__all", q:"", open:new Set() };

/* ---------- utilitários ---------- */
const num = v => { if(v==null||v==="") return 0; if(typeof v==="number") return isFinite(v)?v:0;
  const s=String(v).replace(/[R$\s]/g,""); const n = s.includes(",") ? parseFloat(s.replace(/\./g,"").replace(",",".")) : parseFloat(s); return isFinite(n)?n:0; };
const prob = v => { if(v==null||v==="") return null; if(typeof v==="number") return v<=1? Math.round(v*100) : Math.round(v);
  const s=String(v), n=parseFloat(s.replace("%","").replace(",",".")); if(!isFinite(n)) return null; return (n<=1 && !s.includes("%")) ? Math.round(n*100) : Math.round(n); };
const parseDate = v => { if(v==null||v==="") return null; if(v instanceof Date) return new Date(v.getFullYear(),v.getMonth(),v.getDate());
  if(typeof v==="number"){ const d=XLSX.SSF.parse_date_code(v); return d? new Date(d.y,d.m-1,d.d):null; }
  let m=String(v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); if(m) return new Date(+m[3],+m[2]-1,+m[1]);
  m=String(v).match(/(\d{4})-(\d{2})-(\d{2})/); return m? new Date(+m[1],+m[2]-1,+m[3]) : null; };
const iso = d => d ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}` : null;
const execName = s => String(s||"Sem responsável").replace(/\s*\(.*?\)\s*$/,"").trim();
const title = s => String(s||"").toLowerCase().replace(/(^|\s|\/|-)(\p{L})/gu,(a,b,c)=>b+c.toUpperCase()).replace(/\b(Ltda|Sa|Me|Epp|De|Da|Do|Dos|Das|E|Em)\b/g,w=>({Ltda:"Ltda",Sa:"SA",Me:"ME",Epp:"EPP"}[w]||w.toLowerCase()));
const nm = e => title(execName(e));
const brl = v => { const a=Math.abs(v);
  if(a>=1e6) return "R$ "+(v/1e6).toLocaleString("pt-BR",{maximumFractionDigits:2})+" mi";
  if(a>=1e3) return "R$ "+(v/1e3).toLocaleString("pt-BR",{maximumFractionDigits:1})+" mil";
  return "R$ "+v.toLocaleString("pt-BR",{maximumFractionDigits:0}); };
const full = v => v.toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
const sum = (a,k) => a.reduce((s,o)=>s+o[k],0);
const agg = a => ({n:a.length, contas:new Set(a.map(o=>o.acc)).size, ad:sum(a,"ad"), im:sum(a,"im"), mrr:sum(a,"mrr")});
const esc = s => String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const dt = s => new Date(s).toLocaleString("pt-BR",{day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"});
const say = (t,ok) => { const m=$("msg"); m.textContent=t; m.className="msg "+(ok?"ok":"err"); if(ok) setTimeout(()=>{ if(m.textContent===t) m.className="msg"; },6000); };

/* linha no formato do banco -> objeto usado nas telas */
function toView(r){
  const d = parseDate(r.data_prevista);
  const o = { code:String(r.codigo??""), desc:r.descricao||"", acc:String(r.conta||"Sem conta").trim(), tipo:r.tipo||"",
    etapa:String(r.etapa||"Sem etapa").trim(), p: r.probabilidade==null? null : Number(r.probabilidade), d,
    mi: d? d.getFullYear()*12+d.getMonth() : null, exec:r.executivo, cargaEm:r.carga_em,
    saas:Number(r.valor_saas)||0, sms:Number(r.valor_sms)||0, ad:Number(r.valor_cdu)||0, im:Number(r.valor_servicos)||0 };
  o.mrr = o.saas + o.sms;
  o.fc = o.mi===CUR && o.p!=null && o.p>=60;
  o.pp = o.mi!=null && o.mi>=CUR && o.mi<=CUR+2;
  return o;
}
/* linha da planilha -> formato do banco */
function fromSheet(r){
  return { codigo:String(r["Código"]??"").trim(), descricao:r["Descrição"]? String(r["Descrição"]):null,
    conta:String(r["Nome da conta"]||"").trim()||"Sem conta", tipo:r["Tipo de oportunidade"]||null, etapa:r["Etapa"]? String(r["Etapa"]).trim():null,
    probabilidade:prob(r["Probabilidade"]), data_prevista:iso(parseDate(r["Data prevista (dias)"])),
    executivo:String(r["Responsável"]||"").trim()||"Sem responsável",
    valor_saas:num(r["Valor SAAS"]), valor_cdu:num(r["Valor CDU/Adesão"]), valor_sms:num(r["Valor SMS"]), valor_servicos:num(r["Valor Serviços Não Recorrentes"]) };
}

/* ---------- autenticação ---------- */
async function boot(){
  if(CFG.LOGO){ const i=$("logo"); i.src=CFG.LOGO; i.classList.add("on"); }
  const { data:{ session } } = await db.auth.getSession();
  session ? entrar() : mostrarLogin();
  db.auth.onAuthStateChange((ev)=>{ if(ev==="SIGNED_OUT") mostrarLogin(); });
}
function mostrarLogin(){ $("app").classList.add("hidden"); $("login").classList.remove("hidden"); }
async function entrar(){
  $("login").classList.add("hidden"); $("app").classList.remove("hidden");
  const { data } = await db.rpc("meu_papel"); papel = data || "usuario";
  $("limparBox").classList.toggle("hidden", papel!=="dono");
  $("histAviso").textContent = papel==="dono" ? "Desfazer uma carga faz os executivos dela voltarem para a carga anterior." : "Registro de todas as cargas feitas no painel.";
  await carregar();
}

$("loginForm").addEventListener("submit", async e=>{
  e.preventDefault(); $("loginErr").textContent=""; $("btnEntrar").disabled=true;
  const { error } = await db.auth.signInWithPassword({ email:$("email").value.trim(), password:$("senha").value });
  $("btnEntrar").disabled=false;
  if(error){ $("loginErr").textContent = /invalid/i.test(error.message) ? "E-mail ou senha incorretos." : "Não foi possível entrar: "+error.message; return; }
  $("senha").value=""; entrar();
});
$("btnSair").addEventListener("click", ()=>db.auth.signOut());

/* ---------- leitura do banco ---------- */
async function carregar(){
  $("source").textContent="Carregando…";
  const all=[]; let from=0;
  while(true){
    const { data, error } = await db.from("oportunidades_atuais").select("*").order("executivo").order("id").range(from, from+999);
    if(error){ say(erroAmigavel(error), false); $("source").textContent=""; return; }
    all.push(...data); if(data.length<1000) break; from+=1000;
  }
  rows = all.map(toView);
  const { data:uc } = await db.from("cargas").select("criada_em,criada_por").eq("desfeita",false).order("criada_em",{ascending:false}).limit(1);
  ultimaCarga = uc && uc[0] || null;
  if(state.exec!=="__all" && !rows.some(o=>o.exec===state.exec)) state.exec="__all";
  render();
}
function erroAmigavel(err){
  const m = err && (err.message||String(err)) || "";
  if(/Somente o responsável/i.test(m)) return m;
  if(/Acesso negado|permission|row-level/i.test(m)) return "Seu usuário não tem permissão. Confira se o e-mail está na tabela de administradores.";
  if(/Failed to fetch|NetworkError/i.test(m)) return "Sem conexão com o banco de dados. Verifique a internet e o arquivo config.js.";
  return "Erro: "+m;
}

/* ---------- upload ---------- */
$("file").addEventListener("change", async e=>{
  const files=[...e.target.files]; e.target.value=""; if(!files.length) return;
  try{
    const porArquivo=[]; let linhas=[];
    for(const f of files){
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf,{type:"array",cellDates:true});
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:null});
      if(!raw.length) throw new Error(`A planilha ${f.name} está vazia.`);
      const miss = REQUIRED.filter(c=>!Object.keys(raw[0]).includes(c));
      if(miss.length) throw new Error(`A planilha ${f.name} não tem as colunas: ${miss.join(", ")}. Use o mesmo modelo exportado do CRM.`);
      const l = raw.map(fromSheet).filter(x=>x.codigo);
      porArquivo.push({ nome:f.name, qtd:l.length }); linhas.push(...l);
    }
    // mesma oportunidade em mais de um arquivo: vale a última
    const porCodigo=new Map(); linhas.forEach(l=>porCodigo.set(l.codigo,l)); const dup=linhas.length-porCodigo.size; linhas=[...porCodigo.values()];
    const execs=[...new Set(linhas.map(l=>l.executivo))].sort();
    pendente = { arquivos:files.map(f=>f.name), linhas, execs };
    $("cargaBody").innerHTML =
      `<ul class="files">${porArquivo.map(a=>`<li>${esc(a.nome)}: ${a.qtd} oportunidades</li>`).join("")}</ul>`+
      (dup?`<p style="font-size:13px;color:var(--muted)">${dup} oportunidade(s) repetida(s) entre arquivos foram consideradas uma vez só.</p>`:"")+
      `<div class="tbl"><table><thead><tr><th>Executivo</th><th class="n">Hoje no painel</th><th class="n">Depois da carga</th><th class="n">Contas</th></tr></thead><tbody>`+
      execs.map(x=>{ const hoje=rows.filter(o=>o.exec===x).length, nov=linhas.filter(l=>l.executivo===x);
        return `<tr><td>${esc(nm(x))}${hoje?"":' <span class="badge clean">novo</span>'}</td><td class="n">${hoje}</td><td class="n"><b>${nov.length}</b></td><td class="n">${new Set(nov.map(l=>l.conta)).size}</td></tr>`}).join("")+
      `</tbody></table></div>`;
    $("dlgCarga").showModal();
  }catch(err){ say(err.message||"Não foi possível ler o arquivo. Envie um .xls ou .xlsx.", false); }
});
$("cargaCancelar").addEventListener("click", ()=>{ pendente=null; $("dlgCarga").close(); });
$("cargaConfirmar").addEventListener("click", async ()=>{
  if(!pendente) return; const btn=$("cargaConfirmar"); btn.disabled=true; btn.textContent="Gravando…";
  const { error } = await db.rpc("registrar_carga", { p_arquivos:pendente.arquivos, p_executivos:pendente.execs, p_linhas:pendente.linhas });
  btn.disabled=false; btn.textContent="Confirmar carga"; $("dlgCarga").close();
  if(error){ say(erroAmigavel(error), false); return; }
  say(`Carga gravada: ${pendente.linhas.length} oportunidades de ${pendente.execs.length} executivo(s).`, true);
  pendente=null; await carregar();
});
document.querySelector(".upload").addEventListener("keydown",e=>{ if(e.key==="Enter"||e.key===" "){ e.preventDefault(); $("file").click(); } });

/* ---------- histórico / desfazer / limpar ---------- */
$("btnHist").addEventListener("click", abrirHistorico);
$("histFechar").addEventListener("click", ()=>$("dlgHist").close());
async function abrirHistorico(){
  $("histTbl").innerHTML=`<tr><td>Carregando…</td></tr>`; $("dlgHist").showModal();
  const { data, error } = await db.from("cargas").select("id,criada_em,criada_por,arquivos,tipo,desfeita,carga_executivos(executivo,qtd)").order("criada_em",{ascending:false}).limit(30);
  if(error){ $("histTbl").innerHTML=`<tr><td>${esc(erroAmigavel(error))}</td></tr>`; return; }
  $("histTbl").innerHTML = `<thead><tr><th>Data</th><th>Por</th><th>Executivos</th><th>Arquivos</th><th></th></tr></thead><tbody>`+
    (data.length? data.map(c=>`<tr style="${c.desfeita?"opacity:.55":""}"><td style="white-space:nowrap">${dt(c.criada_em)}</td><td>${esc((c.criada_por||"").split("@")[0])}</td>
      <td>${c.carga_executivos.map(x=>`${esc(nm(x.executivo))} (${x.qtd})`).join("<br>")}</td>
      <td>${c.tipo==="limpeza"?'<span class="badge clean">limpeza</span>':esc((c.arquivos||[]).join(", "))}</td>
      <td class="n">${c.desfeita?'<span class="badge undo">desfeita</span>':(papel==="dono"?`<button data-undo="${c.id}">Desfazer</button>`:"")}</td></tr>`).join("")
      : `<tr><td colspan="5" style="color:var(--muted)">Nenhuma carga ainda.</td></tr>`)+"</tbody>";
  const execs=[...new Set(rows.map(o=>o.exec))].sort();
  $("limparSel").innerHTML = execs.length? execs.map(x=>`<option value="${esc(x)}">${esc(nm(x))}</option>`).join("") : `<option value="">Nenhum executivo</option>`;
}
$("histTbl").addEventListener("click", async e=>{
  const b=e.target.closest("[data-undo]"); if(!b) return;
  if(!confirm("Desfazer esta carga? Os executivos dela voltam para a carga anterior.")) return;
  b.disabled=true; const { error } = await db.rpc("desfazer_carga",{ p_carga:Number(b.dataset.undo) });
  if(error){ alert(erroAmigavel(error)); b.disabled=false; return; }
  await carregar(); abrirHistorico(); say("Carga desfeita.", true);
});
$("btnLimpar").addEventListener("click", async ()=>{
  const x=$("limparSel").value; if(!x) return;
  if(!confirm(`Limpar a carteira de ${nm(x)}? Ele deixa de aparecer no painel. Dá para desfazer pelo histórico.`)) return;
  const { error } = await db.rpc("registrar_carga",{ p_arquivos:[], p_executivos:[x], p_linhas:[] });
  if(error){ alert(erroAmigavel(error)); return; }
  await carregar(); abrirHistorico(); say(`Carteira de ${nm(x)} limpa.`, true);
});

/* ---------- painel ---------- */
const HZ = {
  fc:{name:"Forecast", rule:"Mês atual, probabilidade de 60% ou mais", f:o=>o.fc},
  pp:{name:"Pipeline", rule:"Mês atual e os dois seguintes", f:o=>o.pp},
  tt:{name:"Total", rule:"Todas as oportunidades em andamento", f:()=>true}
};
const scoped = () => rows.filter(o=>state.exec==="__all"||o.exec===state.exec);
const inHz = () => scoped().filter(HZ[state.hz].f);
const topEtapa = v => v.map(o=>o.etapa).sort((x,y)=>y.localeCompare(x,"pt-BR",{numeric:true}))[0];

function render(){
  const base=scoped();
  $("source").textContent = (ultimaCarga? `Última carga em ${dt(ultimaCarga.criada_em)} por ${(ultimaCarga.criada_por||"").split("@")[0]}` : "Nenhuma carga ainda") + ` · referência: ${MES[TODAY.getMonth()]}/${TODAY.getFullYear()}`;
  const execs=[...new Set(rows.map(o=>o.exec))].sort((a,b)=>nm(a).localeCompare(nm(b),"pt-BR"));
  $("exec").innerHTML=`<option value="__all">Todos os executivos (${execs.length})</option>`+execs.map(e=>`<option value="${esc(e)}" ${e===state.exec?"selected":""}>${esc(nm(e))}</option>`).join("");

  if(!rows.length){
    $("execBoards").innerHTML=`<div class="empty"><b>Nenhuma oportunidade no painel</b>Clique em "Carregar planilhas" e envie a exportação do CRM. Pode ser um arquivo por executivo ou um único arquivo com vários.</div>`;
    $("dash").classList.add("hidden"); return;
  }
  $("dash").classList.remove("hidden");

  // quadro por executivo
  const rk=v=>v.some(o=>o.fc)?0:v.some(o=>o.pp)?1:2;
  $("execBoards").innerHTML=[...new Set(base.map(o=>o.exec))].sort((a,b)=>nm(a).localeCompare(nm(b),"pt-BR")).map(e=>{
    const a=base.filter(o=>o.exec===e), m={}; a.forEach(o=>{(m[o.acc]=m[o.acc]||[]).push(o)});
    const list=Object.entries(m).sort((x,y)=>rk(x[1])-rk(y[1])||x[0].localeCompare(y[0],"pt-BR"));
    return `<div class="board"><h2>Executivo: ${esc(nm(e))}<small>${list.length} contas em aberto · ${a.length} oportunidades</small></h2>
      <div class="meta">Atualizado em ${a[0].cargaEm? dt(a[0].cargaEm):"—"}</div>
      <div class="leg"><span><i style="background:var(--fc)"></i>No forecast</span><span><i style="background:var(--pp)"></i>No pipeline</span><span><i style="background:var(--tt)"></i>Demais em andamento</span></div>
      <div class="tiles">${list.map(([k,v])=>`<button class="tile ${["fc","pp","tt"][rk(v)]}" data-acc="${esc(k)}"><span class="an">${esc(title(k))}</span><span class="am">${v.length} ${v.length>1?"oportunidades":"oportunidade"} · ${esc(topEtapa(v))}</span></button>`).join("")}</div></div>`}).join("");

  // horizontes
  $("hz").innerHTML=Object.entries(HZ).map(([k,h])=>{ const s=agg(base.filter(h.f)); const won=base.filter(o=>h.f(o)&&o.p===100).length;
    return `<button class="hz ${k}" data-h="${k}" aria-pressed="${state.hz===k}">
      <div class="name">${h.name}${k==="fc"?'<span class="tag">Compromisso do mês</span>':""}</div><div class="rule">${h.rule}</div>
      <div class="counts"><b>${s.n}</b>oportunidades &nbsp; <b>${s.contas}</b>contas${k==="fc"&&won?` &nbsp;· ${won} ganha(s)`:""}</div>
      <div class="vals"><div class="vrow"><span><i class="dot d-ad"></i>Adesão (CDU)</span><b class="v-ad" title="${full(s.ad)}">${brl(s.ad)}</b></div>
      <div class="vrow"><span><i class="dot d-im"></i>Implantação (serviços)</span><b class="v-im" title="${full(s.im)}">${brl(s.im)}</b></div>
      <div class="vrow"><span><i class="dot d-rr"></i>RR (mensal)</span><b class="v-rr" title="${full(s.mrr)}">${brl(s.mrr)}</b></div></div></button>`}).join("");

  const cur=inHz(), hn=HZ[state.hz].name;

  // funil
  $("funnelDesc").textContent=`${hn}: ${cur.length} oportunidades por etapa.`;
  const etapas=[...new Set(rows.map(o=>o.etapa))].sort((a,b)=>a.localeCompare(b,"pt-BR",{numeric:true}));
  const fmax=Math.max(1,...etapas.map(e=>{const a=cur.filter(o=>o.etapa===e);return sum(a,"ad")+sum(a,"im")+sum(a,"mrr")}));
  $("funnel").innerHTML=`<div class="row head"><div>Etapa</div><div></div><div>Adesão + Impl.<br>RR (mensal)</div></div>`+etapas.map(e=>{
    const a=cur.filter(o=>o.etapa===e), ad=sum(a,"ad"), im=sum(a,"im"), rr=sum(a,"mrr");
    return `<div class="row" tabindex="0"><div>${esc(e)} <span style="color:var(--muted)">(${a.length})</span></div>
    <div class="track"><i class="c-ad" style="width:${ad/fmax*100}%"></i><i class="c-im" style="width:${im/fmax*100}%"></i><i class="c-mr" style="width:${rr/fmax*100}%"></i></div>
    <div style="text-align:right">${brl(ad+im)}<br><small class="v-rr">RR ${brl(rr)}</small></div>
    <div class="tip" role="tooltip"><b>${esc(e)}</b><span>${a.length} ${a.length===1?"oportunidade":"oportunidades"} · ${new Set(a.map(o=>o.acc)).size} contas</span>
      <div class="tl"><i class="c-ad"></i>Adesão (CDU)<em>${full(ad)}</em></div>
      <div class="tl"><i class="c-im"></i>Implantação (serviços)<em>${full(im)}</em></div>
      <div class="tl sep">Adesão + Implantação<em>${full(ad+im)}</em></div>
      <div class="tl"><i class="c-mr"></i>RR (mensalidade)<em>${full(rr)}</em></div></div></div>`}).join("");

  // alertas
  const al=[], grp=a=>{const m={};a.forEach(o=>{(m[o.acc]=m[o.acc]||[]).push(o)});return Object.entries(m).map(([k,v])=>`${esc(title(k))} (${v.length})`).join(", ")};
  const a1=base.filter(o=>o.mi===CUR&&!o.fc); if(a1.length) al.push(["Fecham este mês, mas estão fora do forecast (abaixo de 60%)", `${grp(a1)} · ${brl(sum(a1,"ad")+sum(a1,"im"))} + RR ${brl(sum(a1,"mrr"))}`]);
  const a2=base.filter(o=>/^\s*[67]\./.test(o.etapa)&&o.p!=null&&o.p<=30); if(a2.length) al.push(["Negociação ou fechamento com probabilidade de 30% ou menos", grp(a2)]);
  const a3=base.filter(o=>o.p==null); if(a3.length) al.push(["Sem probabilidade preenchida", grp(a3)]);
  const a4=base.filter(o=>o.mi!=null&&o.mi<CUR); if(a4.length) al.push(["Data prevista já vencida", grp(a4)]);
  const a5=base.filter(o=>o.mi==null); if(a5.length) al.push(["Sem data prevista", grp(a5)]);
  const a6=base.filter(o=>o.ad+o.im+o.mrr===0); if(a6.length) al.push(["Oportunidades sem valor", grp(a6)]);
  $("alerts").innerHTML= al.length? al.map(([t,d])=>`<li><b>${t}</b>${d}</li>`).join("") : `<li class="ok"><b>Nada a corrigir</b>Todas as oportunidades estão consistentes.</li>`;

  // executivos: um bloco por executivo com os três horizontes
  const blk=(k,a,ref,refLabel)=>{ const s=agg(a), oneTime=s.ad+s.im, refV=ref? sum(ref,"ad")+sum(ref,"im"):0, pct=refV? Math.round(oneTime/refV*100):null;
    return `<div class="xh ${k}"><div class="xt"><b>${HZ[k].name}</b><em>${s.n} opp · ${s.contas} contas</em></div>
      <div class="xv"><span><i class="dot d-ad"></i>Adesão</span><strong class="v-ad" title="${full(s.ad)}">${brl(s.ad)}</strong></div>
      <div class="xv"><span><i class="dot d-im"></i>Implantação</span><strong class="v-im" title="${full(s.im)}">${brl(s.im)}</strong></div>
      <div class="xv"><span><i class="dot d-rr"></i>RR mensal</span><strong class="v-rr" title="${full(s.mrr)}">${brl(s.mrr)}</strong></div>
      ${ref?`<div class="xshare">${pct==null?"—":pct+"%"} ${refLabel}<div class="xbar"><i style="width:${Math.min(100,pct||0)}%"></i></div></div>`:""}</div>`; };
  const ini=e=>{const p=execName(e).split(/\s+/).filter(w=>w.length>2);return ((p[0]||"?")[0]+((p[1]||"")[0]||"")).toUpperCase()};
  const xrow=(label,sub,a,cls,avatar)=>{ const pp=a.filter(o=>o.pp);
    return `<div class="xrow ${cls||""}"><div class="xwho"><span class="av">${avatar}</span><div><b>${esc(label)}</b><small>${sub}</small></div></div>
      ${blk("fc",a.filter(o=>o.fc),pp,"do pipeline (adesão + impl.)")}${blk("pp",pp,a,"do total (adesão + impl.)")}${blk("tt",a)}</div>`; };
  const ordem=[...execs].sort((x,y)=>{const f=e=>{const a=rows.filter(o=>o.exec===e);return [sum(a.filter(o=>o.fc),"ad")+sum(a.filter(o=>o.fc),"im"), sum(a.filter(o=>o.pp),"ad")+sum(a.filter(o=>o.pp),"im")]};const A=f(x),B=f(y);return B[0]-A[0]||B[1]-A[1]});
  $("execList").innerHTML=(execs.length>1? xrow(`Time (${execs.length} executivos)`, `${new Set(rows.map(o=>o.acc)).size} contas · ${rows.length} oportunidades`, rows, "team", "∑") : "")+
    ordem.map(e=>{const a=rows.filter(o=>o.exec===e);return xrow(nm(e), `${new Set(a.map(o=>o.acc)).size} contas · ${a.length} oportunidades`, a, "", ini(e))}).join("");

  // contas
  const q=state.q.toLowerCase(), byAcc={}; cur.forEach(o=>{(byAcc[o.acc]=byAcc[o.acc]||[]).push(o)});
  const peso=v=>sum(v,"ad")+sum(v,"im")+sum(v,"mrr");
  const accs=Object.entries(byAcc).filter(([k,v])=>!q||k.toLowerCase().includes(q)||v.some(o=>o.code.includes(q))).sort((a,b)=>peso(b[1])-peso(a[1]));
  $("accTbl").innerHTML=`<thead><tr><th>Conta</th><th class="n">Opp</th><th>Etapa mais avançada</th><th>Fechamento</th><th class="n h-ad">Adesão</th><th class="n h-im">Implantação</th><th class="n h-rr">RR</th></tr></thead><tbody>`+
   (accs.length? accs.map(([k,v])=>{ const dmin=v.filter(o=>o.d).map(o=>o.d).sort((a,b)=>a-b)[0], open=state.open.has(k);
     let h=`<tr class="acc" data-a="${esc(k)}" tabindex="0" aria-expanded="${open}"><td><b style="font-weight:600">${esc(title(k))}</b></td><td class="n">${v.length}</td><td>${esc(topEtapa(v))}</td><td>${dmin?dmin.toLocaleDateString("pt-BR"):"—"}</td><td class="n v-ad">${brl(sum(v,"ad"))}</td><td class="n v-im">${brl(sum(v,"im"))}</td><td class="n v-rr">${brl(sum(v,"mrr"))}</td></tr>`;
     if(open) h+=v.map(o=>`<tr class="det"><td>${esc(o.code)} · ${esc(o.desc)}<br>${esc(o.tipo)} · ${esc(nm(o.exec))}</td><td class="n"><span class="pill ${o.fc?"hi":""}">${o.p==null?"s/ prob.":o.p+"%"}</span></td><td>${esc(o.etapa)}</td><td>${o.d?o.d.toLocaleDateString("pt-BR"):"—"}</td><td class="n">${full(o.ad)}</td><td class="n">${full(o.im)}</td><td class="n">${full(o.mrr)}</td></tr>`).join("");
     return h; }).join("") : `<tr><td colspan="7" style="color:var(--muted)">Nenhuma conta neste horizonte${q?" com essa busca":""}.</td></tr>`)+"</tbody>";
}

/* ---------- interações ---------- */
$("exec").addEventListener("change",e=>{ state.exec=e.target.value; render(); });
$("q").addEventListener("input",e=>{ state.q=e.target.value; render(); $("q").focus(); });
document.addEventListener("click",e=>{
  const h=e.target.closest("[data-h]"); if(h){ state.hz=h.dataset.h; render(); return; }
  const t=e.target.closest(".tile"); if(t){ const k=t.dataset.acc; state.hz="tt"; state.q=""; $("q").value=""; state.open.add(k); render();
    const r=[...document.querySelectorAll("tr.acc")].find(x=>x.dataset.a===k); if(r) r.scrollIntoView({behavior:"smooth",block:"center"}); return; }
  const a=e.target.closest("tr.acc"); if(a){ const k=a.dataset.a; state.open.has(k)?state.open.delete(k):state.open.add(k); render(); }
});
document.addEventListener("keydown",e=>{ const a=e.target.closest&&e.target.closest("tr.acc"); if(a&&(e.key==="Enter"||e.key===" ")){ e.preventDefault(); a.click(); } });

boot();
