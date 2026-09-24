/* Dashboard de Oportunidades — GitHub Pages + Supabase */
const CFG = window.APP_CONFIG || {};
const db = window.supabase.createClient(CFG.SUPABASE_URL, CFG.SUPABASE_KEY);

const REQUIRED = ["Código","Nome da conta","Tipo de oportunidade","Etapa","Probabilidade","Data prevista (dias)","Responsável","Valor SAAS","Valor CDU/Adesão","Valor SMS","Valor Serviços Não Recorrentes"];
const OPCIONAIS = ["Descrição","Valor Serviços Recorrentes","Situação"];
const ALIAS = { "status":"Situação", "valor scc":"Valor Serviços Recorrentes", "scc":"Valor Serviços Recorrentes", "valor saas":"Valor SAAS" };
const normH = s => String(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/\bde\b/g," ").replace(/[^a-z0-9%()/]+/g," ").replace(/\s+/g," ").trim();
const CANON = {}; REQUIRED.concat(OPCIONAIS).forEach(c=>CANON[normH(c)]=c); Object.entries(ALIAS).forEach(([k,v])=>CANON[normH(k)]=v);
const canonRow = r => { const o={}; for(const k in r){ const c=CANON[normH(k)]||k; if(!(c in o) || o[c]==null) o[c]=r[k]; } return o; };
const MES = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
const TODAY = new Date();
const CUR = TODAY.getFullYear()*12 + TODAY.getMonth();
const $ = id => document.getElementById(id);

let rows = [], ganhas = [], metas = [], metasOn = true, ultimaCarga = null, pendente = null, papel = "usuario", meuT = null;
const DOM_T = CFG.DOMINIO_LOGIN_T || "vendedor.example.com";
const codigoT = e => { const m=String(e||"").match(/\(\s*(T\d+)\s*\)/i); return m? m[1].toUpperCase() : null; };
let state = { pv:"ad", hz:"tt", exec:"__all", q:"", open:new Set(), probs:new Set(), meses:new Set(), dtOpen:false, exp:new Set() };
const pKey = o => o.p==null? "s" : String(o.p);
const pOk0 = o => !state.probs.size || state.probs.has(pKey(o));
/* filtro de data prevista: ano > trimestre > mês */
const MESL = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
const mKey = o => o.mi==null? "sem" : `${Math.floor(o.mi/12)}-${String(o.mi%12+1).padStart(2,"0")}`;
const dOk = o => !state.meses.size || state.meses.has(mKey(o));
const pOk = o => pOk0(o) && dOk(o);

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
const agg = a => ({n:a.length, contas:new Set(a.map(o=>o.acc)).size, ad:sum(a,"ad"), im:sum(a,"im"), mrr:sum(a,"mrr"), scc:sum(a,"scc")});
const esc = s => String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const dt = s => new Date(s).toLocaleString("pt-BR",{day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"});
const say = (t,ok) => { const m=$("msg"); m.textContent=t; m.className="msg "+(ok?"ok":"err"); if(ok) setTimeout(()=>{ if(m.textContent===t) m.className="msg"; },6000); };

/* linha no formato do banco -> objeto usado nas telas */
function toView(r){
  const d = parseDate(r.data_prevista);
  const o = { code:String(r.codigo??""), desc:r.descricao||"", acc:String(r.conta||"Sem conta").trim(), tipo:r.tipo||"",
    etapa:String(r.etapa||"Sem etapa").trim(), p: r.probabilidade==null? null : Number(r.probabilidade), d,
    mi: d? d.getFullYear()*12+d.getMonth() : null, exec:r.executivo, cargaEm:r.carga_em,
    saas:Number(r.valor_saas)||0, sms:Number(r.valor_sms)||0, ad:Number(r.valor_cdu)||0, im:Number(r.valor_servicos)||0, scc:Number(r.valor_scc)||0 };
  o.mrr = o.saas + o.sms;
  o.fc = o.mi===CUR;   // forecast = data prevista no mês atual (a probabilidade é filtrada à parte)
  o.pp = o.mi!=null && o.mi>=CUR && o.mi<=CUR+2;
  return o;
}
/* linha da planilha -> formato do banco */
function fromSheet(r){
  return { codigo:String(r["Código"]??"").trim(), descricao:r["Descrição"]? String(r["Descrição"]):null,
    conta:String(r["Nome da conta"]||"").trim()||"Sem conta", tipo:r["Tipo de oportunidade"]||null, etapa:r["Etapa"]? String(r["Etapa"]).trim():null,
    probabilidade:prob(r["Probabilidade"]), data_prevista:iso(parseDate(r["Data prevista (dias)"])),
    executivo:String(r["Responsável"]||"").trim()||"Sem responsável",
    valor_saas:num(r["Valor SAAS"]), valor_cdu:num(r["Valor CDU/Adesão"]), valor_sms:num(r["Valor SMS"]), valor_servicos:num(r["Valor Serviços Não Recorrentes"]), valor_scc:num(r["Valor Serviços Recorrentes"]),
    situacao: normH(r["Situação"]||"").toUpperCase() };
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
  const { data } = await db.rpc("meu_perfil");
  papel = (data && data.papel) || "usuario"; meuT = data && data.codigo_t || null;
  const vend = papel==="vendedor";
  $("limparBox").classList.toggle("hidden", papel!=="dono");
  $("btnHist").classList.toggle("hidden", vend);
  $("btnMetas").classList.toggle("hidden", vend);
  $("exec").classList.toggle("hidden", vend);
  $("histAviso").textContent = papel==="dono" ? "Desfazer uma carga faz os executivos dela voltarem para a carga anterior." : "Registro de todas as cargas feitas no painel.";
  if(!data) say("Seu usuário ainda não foi liberado no painel. Peça ao responsável para cadastrar seu acesso.", false);
  await carregar();
  // primeiro acesso: obriga a trocar a senha provisória
  const { data:{ user } } = await db.auth.getUser();
  if(user && !(user.user_metadata && user.user_metadata.senha_definida)) abrirSenha(true);
}

/* ---------- troca de senha ---------- */
let senhaObrigatoria=false;
function abrirSenha(obrigatoria){
  senhaObrigatoria=obrigatoria;
  $("senhaTitulo").textContent = obrigatoria? "Crie a sua senha" : "Alterar senha";
  $("senhaTexto").textContent = obrigatoria? "Este é o seu primeiro acesso. Troque a senha provisória por uma senha só sua." : "Escolha uma nova senha para entrar no painel.";
  $("senhaCancelar").classList.toggle("hidden", obrigatoria);
  $("senha1").value=""; $("senha2").value=""; $("senhaErr").textContent="";
  $("dlgSenha").showModal(); $("senha1").focus();
}
$("btnSenha").addEventListener("click", ()=>abrirSenha(false));
$("senhaCancelar").addEventListener("click", ()=>$("dlgSenha").close());
$("dlgSenha").addEventListener("cancel", e=>{ if(senhaObrigatoria) e.preventDefault(); });
$("senhaForm").addEventListener("submit", async e=>{
  e.preventDefault(); const a=$("senha1").value, b=$("senha2").value;
  if(a.length<8){ $("senhaErr").textContent="A senha precisa ter pelo menos 8 caracteres."; return; }
  if(a!==b){ $("senhaErr").textContent="As duas senhas não são iguais."; return; }
  $("senhaSalvar").disabled=true;
  const { error } = await db.auth.updateUser({ password:a, data:{ senha_definida:true } });
  $("senhaSalvar").disabled=false;
  if(error){ $("senhaErr").textContent = /different/i.test(error.message) ? "A nova senha precisa ser diferente da atual." : /weak|short|characters/i.test(error.message) ? "Senha fraca. Use mais caracteres, misturando letras e números." : "Não foi possível salvar: "+error.message; return; }
  senhaObrigatoria=false; $("dlgSenha").close(); say("Senha alterada com sucesso.", true);
});

$("loginForm").addEventListener("submit", async e=>{
  e.preventDefault(); $("loginErr").textContent=""; $("btnEntrar").disabled=true;
  let login=$("email").value.trim(); if(/^t\d+$/i.test(login)) login=login.toLowerCase()+"@"+DOM_T;
  const { error } = await db.auth.signInWithPassword({ email:login, password:$("senha").value });
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
  rows = all.map(toView).filter(o=>o.p!==100);   // 100% = ganha, não está mais em andamento
  metasOn = true;
  const g=[]; from=0;
  while(true){
    const { data, error } = await db.from("oportunidades_ganhas").select("*").order("id").range(from, from+999);
    if(error){ metasOn=false; break; }
    g.push(...data); if(data.length<1000) break; from+=1000;
  }
  ganhas = g.map(r=>{ const o=toView(r); if(o.mi==null && r.carga_em){ const d=new Date(r.carga_em); o.mi=d.getFullYear()*12+d.getMonth(); } return o; });
  if(metasOn){ const { data, error } = await db.from("metas").select("*"); if(error) metasOn=false; else metas=data||[]; }
  const { data:uc } = await db.from("cargas").select("criada_em,criada_por").eq("desfeita",false).order("criada_em",{ascending:false}).limit(1);
  ultimaCarga = uc && uc[0] || null;
  if(state.exec!=="__all" && !rows.some(o=>o.exec===state.exec)) state.exec="__all";
  render();
}
function erroAmigavel(err){
  const m = err && (err.message||String(err)) || "";
  if(/Somente o responsável/i.test(m)) return m;
  if(/function .*meu_perfil|does not exist/i.test(m)) return "O banco ainda não recebeu a última atualização (script alteracao-vendedor.sql).";
  if(/Acesso negado|permission|row-level/i.test(m)) return "Seu usuário não tem permissão. Confira se o e-mail está na tabela de administradores.";
  if(/Failed to fetch|NetworkError/i.test(m)) return "Sem conexão com o banco de dados. Verifique a internet e o arquivo config.js.";
  return "Erro: "+m;
}

/* ---------- upload ---------- */
$("file").addEventListener("change", async e=>{
  const files=[...e.target.files]; e.target.value=""; if(!files.length) return;
  try{
    const porArquivo=[], semScc=[]; let linhas=[];
    for(const f of files){
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf,{type:"array",cellDates:true});
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:null}).map(canonRow);
      if(!raw.length) throw new Error(`A planilha ${f.name} está vazia.`);
      const cols = new Set(raw.flatMap(r=>Object.keys(r))); const miss = REQUIRED.filter(c=>!cols.has(c)); if(!cols.has("Valor Serviços Recorrentes")) semScc.push(f.name);
      if(miss.length) throw new Error(`A planilha ${f.name} não tem as colunas: ${miss.join(", ")}. Use o mesmo modelo exportado do CRM.`);
      const l = raw.map(fromSheet).filter(x=>x.codigo);
      porArquivo.push({ nome:f.name, qtd:l.length }); linhas.push(...l);
    }
    // Situação: descartadas/perdidas são ignoradas; ganhas (WON) contam como 100%
    const IGN=/^(DISCARDED|DESCARTAD[AO]|LOST|PERDID[AO])$/, WON=/^(WON|GANH[AO])$/;
    const nDesc=linhas.filter(l=>IGN.test(l.situacao)).length;
    linhas=linhas.filter(l=>!IGN.test(l.situacao));
    linhas.forEach(l=>{ if(WON.test(l.situacao)) l.probabilidade=100; delete l.situacao; });
    // vendedor: só as oportunidades do próprio código T
    let ignoradas=0;
    if(papel==="vendedor"){ const antes=linhas.length; linhas=linhas.filter(l=>codigoT(l.executivo)===meuT); ignoradas=antes-linhas.length;
      if(!linhas.length) throw new Error(`Nenhuma oportunidade do seu código (${meuT}) nestas planilhas.`); }
    // mesma oportunidade em mais de um arquivo: vale a última
    const porCodigo=new Map(); linhas.forEach(l=>porCodigo.set(l.codigo,l)); const dup=linhas.length-porCodigo.size; linhas=[...porCodigo.values()];
    const execs=[...new Set(linhas.map(l=>l.executivo))].sort();
    const nGanhas = linhas.filter(l=>l.probabilidade===100).length;
    pendente = { arquivos:files.map(f=>f.name), linhas, execs };
    $("cargaBody").innerHTML =
      `<ul class="files">${porArquivo.map(a=>`<li>${esc(a.nome)}: ${a.qtd} oportunidades</li>`).join("")}</ul>`+
      (ignoradas?`<p style="font-size:13px;color:var(--muted)">${ignoradas} oportunidade(s) de outros executivos foram ignoradas. Você só pode atualizar a carteira do código ${meuT}.</p>`:"")+
      (nDesc?`<p style="font-size:13px;color:var(--muted)">${nDesc} oportunidade(s) com situação descartada ou perdida foram ignoradas.</p>`:"")+
      (nGanhas?`<p style="font-size:13px;color:var(--muted)">${nGanhas} oportunidade(s) com 100% serão registradas como <b>ganhas</b> e passam a contar no Realizado das metas.</p>`:"")+
      (semScc.length?`<p style="font-size:13px;color:var(--muted)">Sem a coluna "Valor Serviços Recorrentes" (SCC): ${semScc.map(esc).join(", ")}. O SCC dessas oportunidades fica zerado.</p>`:"")+
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

/* ---------- ajuda do upload ---------- */
const AJUDA = [
  ["Código","Código da oportunidade no CRM (único por oportunidade)"],
  ["Nome da conta","Nome da empresa cliente"],
  ["Tipo de oportunidade","Software ou Serviços"],
  ["Etapa","Etapa do funil, ex.: 5. Propostas"],
  ["Probabilidade","10%, 30%, 60%, 90% ou 100%"],
  ["Data prevista (dias)","Data prevista de fechamento, ex.: 31/10/2026"],
  ["Responsável","Nome do executivo com o código T entre parênteses, ex.: JOSE DA SILVA (T12345)"],
  ["Valor SAAS","Mensalidade SaaS (entra no RR)"],
  ["Valor SMS","Mensalidade SMS (entra no RR)"],
  ["Valor CDU/Adesão","Valor de adesão / CDU (não recorrente)"],
  ["Valor Serviços Não Recorrentes","Valor de implantação / serviços"],
  ["Valor Serviços Recorrentes","SCC: serviços recorrentes mensais (se não vier, fica zerado)"],
  ["Situação","Opcional: WON conta como ganha (100%); DISCARDED e LOST são ignoradas"],
  ["Descrição","Opcional: descrição da oportunidade"]
];
$("ajudaCols").innerHTML = AJUDA.map(([c,d])=>`<tr><td style="white-space:nowrap"><b style="font-weight:600">${esc(c)}</b>${OPCIONAIS.includes(c)?' <span class="badge">opcional</span>':""}</td><td style="color:var(--muted)">${esc(d)}</td></tr>`).join("");
const pularAjuda = () => { try{ return localStorage.getItem("opp-ajuda-upload")==="nao"; }catch(e){ return false; } };
$("btnUpload").addEventListener("click", ()=>{ if(pularAjuda()) $("file").click(); else { $("ajudaNaoMostrar").checked=false; $("dlgAjuda").showModal(); } });
$("verAjuda").addEventListener("click", e=>{ e.preventDefault(); $("ajudaNaoMostrar").checked=false; $("dlgAjuda").showModal(); });
$("ajudaCancelar").addEventListener("click", ()=>$("dlgAjuda").close());
$("ajudaEscolher").addEventListener("click", ()=>{
  if($("ajudaNaoMostrar").checked){ try{ localStorage.setItem("opp-ajuda-upload","nao"); }catch(e){} }
  $("dlgAjuda").close(); $("file").click();
});
$("ajudaModelo").addEventListener("click", ()=>{
  const cab = REQUIRED.concat(["Valor Serviços Recorrentes","Situação","Descrição"]);
  const ex = ["123456","EMPRESA EXEMPLO LTDA","Software","5. Propostas","30%","31/12/2026","JOSE DA SILVA (T12345)",1500,20000,800,35000,1200,"OPEN","Exemplo de oportunidade"];
  const ws = XLSX.utils.aoa_to_sheet([cab, ex]); ws["!cols"]=cab.map(c=>({wch:Math.max(14,c.length+2)}));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Oportunidades");
  XLSX.writeFile(wb, "modelo-oportunidades.xlsx");
});

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

/* ---------- árvore de datas ---------- */
function arvoreDatas(){
  const cont={}; rows.forEach(o=>{ const k=mKey(o); cont[k]=(cont[k]||0)+1; });
  const anos={};
  Object.keys(cont).filter(k=>k!=="sem").sort().forEach(k=>{ const [y,m]=k.split("-").map(Number), q=Math.ceil(m/3);
    ((anos[y]=anos[y]||{})[q]=anos[y][q]||[]).push({k,m,n:cont[k]}); });
  return { anos, sem:cont.sem||0, todos:Object.keys(cont) };
}
function rotuloDatas(t){
  const sel=[...state.meses]; if(!sel.length) return "Todos";
  const has=ks=>ks.every(k=>state.meses.has(k));
  for(const y in t.anos){ const ks=Object.values(t.anos[y]).flat().map(x=>x.k); if(sel.length===ks.length && has(ks)) return y;
    for(const q in t.anos[y]){ const kq=t.anos[y][q].map(x=>x.k); if(sel.length===kq.length && has(kq)) return `${q}º Tri ${y}`; } }
  if(sel.length===1){ if(sel[0]==="sem") return "Sem data"; const [y,m]=sel[0].split("-"); return `${MESL[+m-1]}/${y}`; }
  return `${sel.length} meses`;
}
function htmlDatas(t){
  const ck=(ks,label,extra,node,n)=>{ const c=ks.filter(k=>state.meses.has(k)).length;
    return `<label class="dn ${extra||""}"><input type="checkbox" data-node="${esc(node)}" data-keys="${ks.join(",")}" ${c&&c===ks.length?"checked":""} ${c&&c<ks.length?'data-ind="1"':""}><span>${label}</span>${n!=null?`<em>${n}</em>`:""}</label>`; };
  const tw=(id)=>`<button type="button" class="tw ${state.exp.has(id)?"open":""}" data-exp="${id}" aria-label="Expandir">›</button>`;
  let h=`<div class="drow l0"><span class="tw sp"></span>${ck(t.todos,"Selecionar tudo","","all")}</div>`;
  Object.keys(t.anos).sort().forEach(y=>{ const yk=Object.values(t.anos[y]).flat().map(x=>x.k), yn=Object.values(t.anos[y]).flat().reduce((a,x)=>a+x.n,0);
    h+=`<div class="drow l1">${tw("y"+y)}${ck(yk,y,"b","y"+y,yn)}</div>`;
    if(state.exp.has("y"+y)) Object.keys(t.anos[y]).sort().forEach(q=>{ const ms=t.anos[y][q], qk=ms.map(x=>x.k), qn=ms.reduce((a,x)=>a+x.n,0);
      h+=`<div class="drow l2">${tw("q"+y+q)}${ck(qk,q+"º Tri","","q"+y+q,qn)}</div>`;
      if(state.exp.has("q"+y+q)) ms.forEach(x=>{ h+=`<div class="drow l3"><span class="tw sp"></span>${ck([x.k],MESL[x.m-1],"","m"+x.k,x.n)}</div>`; });
    });
  });
  if(t.sem) h+=`<div class="drow l1"><span class="tw sp"></span>${ck(["sem"],"Sem data prevista","","sem",t.sem)}</div>`;
  return h;
}

/* ---------- painel ---------- */
const HZ = {
  fc:{name:"Forecast", rule:"Data prevista no mês atual", f:o=>o.fc},
  pp:{name:"Pipeline", rule:"Mês atual e os dois seguintes", f:o=>o.pp},
  tt:{name:"Total", rule:"Todas as oportunidades em andamento", f:()=>true}
};
const scoped = () => rows.filter(o=>(state.exec==="__all"||o.exec===state.exec) && pOk(o));
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

  // barra de filtros
  const pvals=[...new Set(rows.map(pKey))].sort((a,b)=>a==="s"?1:b==="s"?-1:a-b);
  const hzAtivo = state.hz!=="tt";
  const tD=arvoreDatas(), dAtivo=state.meses.size>0;
  if(!state.expInit && Object.keys(tD.anos).length){ Object.keys(tD.anos).forEach(y=>state.exp.add("y"+y)); state.expInit=true; }
  $("filters").innerHTML = `<div class="dwrap"><span class="fl">Data prevista</span>
      <button type="button" class="dbtn ${dAtivo?"on":""}" data-dt="1" aria-expanded="${state.dtOpen}">${esc(rotuloDatas(tD))}<i>▾</i></button>
      ${state.dtOpen?`<div class="dpop" role="dialog" aria-label="Filtro de data prevista">${htmlDatas(tD)}</div>`:""}</div>
    <span class="fsep"></span><span class="fl">Probabilidade</span>
    <button class="chip all ${state.probs.size?"":"on"}" data-p="__all" aria-pressed="${!state.probs.size}">Todas</button>`+
    pvals.map(v=>`<button class="chip ${state.probs.has(v)?"on":""}" data-p="${v}" aria-pressed="${state.probs.has(v)}">${v==="s"?"Sem probabilidade":v+"%"}</button>`).join("")+
    `<span class="pcount">${state.probs.size? `<b>${state.probs.size}</b> de ${pvals.length} selecionadas` : `todas as ${pvals.length} faixas`}</span>`+
    (hzAtivo||state.probs.size||dAtivo? `<span class="fsum">Horizonte: <b>${hzAtivo?HZ[state.hz].name:"Total"}</b> · Data: <b>${esc(rotuloDatas(tD))}</b> · Probabilidade: <b>${state.probs.size?[...state.probs].sort((a,b)=>a==="s"?1:b==="s"?-1:a-b).map(v=>v==="s"?"sem prob.":v+"%").join(", "):"todas"}</b></span><button class="chip clear" data-p="__clear">Limpar filtros</button>` : "");

  document.querySelectorAll('.dpop input[data-ind]').forEach(i=>i.indeterminate=true);

  // quadro por executivo (respeita horizonte, probabilidade e data)
  const vis=base.filter(HZ[state.hz].f);
  const rk=v=>v.some(o=>o.fc)?0:v.some(o=>o.pp)?1:2;
  if(!vis.length){ $("execBoards").innerHTML=`<div class="empty"><b>Nenhuma conta com estes filtros</b>Clique de novo no card selecionado ou em "Limpar filtros" para voltar a ver todas.</div>`; }
  else $("execBoards").innerHTML=[...new Set(vis.map(o=>o.exec))].sort((a,b)=>nm(a).localeCompare(nm(b),"pt-BR")).map(e=>{
    const a=vis.filter(o=>o.exec===e), m={}; a.forEach(o=>{(m[o.acc]=m[o.acc]||[]).push(o)});
    const list=Object.entries(m).sort((x,y)=>rk(x[1])-rk(y[1])||x[0].localeCompare(y[0],"pt-BR"));
    return `<div class="board"><h2>Executivo: ${esc(nm(e))}<small>${list.length} contas em aberto · ${a.length} oportunidades</small></h2>
      <div class="meta">Atualizado em ${a[0].cargaEm? dt(a[0].cargaEm):"—"}</div>
      <div class="leg"><span><i style="background:var(--fc)"></i>No forecast</span><span><i style="background:var(--pp)"></i>No pipeline</span><span><i style="background:var(--tt)"></i>Demais em andamento</span></div>
      <div class="tiles">${list.map(([k,v])=>`<button class="tile ${["fc","pp","tt"][rk(v)]}" data-acc="${esc(k)}"><span class="an">${esc(title(k))}</span><span class="am">${v.length} ${v.length>1?"oportunidades":"oportunidade"} · ${esc(topEtapa(v))}</span></button>`).join("")}</div></div>`}).join("");

  // horizontes
  $("hz").innerHTML=Object.entries(HZ).map(([k,h])=>{ const s=agg(base.filter(h.f)); const won=0;
    return `<button class="hz ${k}" data-h="${k}" aria-pressed="${state.hz===k}">
      <div class="name">${h.name}${k==="fc"?'<span class="tag">Compromisso do mês</span>':""}</div><div class="rule">${h.rule}</div>
      <div class="counts"><b>${s.n}</b>oportunidades &nbsp; <b>${s.contas}</b>contas${k==="fc"&&won?` &nbsp;· ${won} ganha(s)`:""}</div>
      <div class="vals"><div class="vrow"><span><i class="dot d-ad"></i>Adesão (CDU)</span><b class="v-ad" title="${full(s.ad)}">${brl(s.ad)}</b></div>
      <div class="vrow"><span><i class="dot d-im"></i>Implantação (serviços)</span><b class="v-im" title="${full(s.im)}">${brl(s.im)}</b></div>
      <div class="vrow"><span><i class="dot d-rr"></i>RR (mensal)</span><b class="v-rr" title="${full(s.mrr)}">${brl(s.mrr)}</b></div>
      <div class="vrow"><span><i class="dot d-scc"></i>SCC (serv. recorrentes)</span><b class="v-scc" title="${full(s.scc)}">${brl(s.scc)}</b></div></div>
      <div class="hint">${state.hz===k? (k==="tt"?"Mostrando todas":"Filtro ativo · clique para ver todas") : "Clique para filtrar"}</div></button>`}).join("");

  renderMetas();

  const cur=inHz(), hn=HZ[state.hz].name;

  // funil
  $("funnelDesc").textContent=`${hn}: ${cur.length} oportunidades por etapa.`;
  const etapas=[...new Set(rows.map(o=>o.etapa))].sort((a,b)=>a.localeCompare(b,"pt-BR",{numeric:true}));
  const fmax=Math.max(1,...etapas.map(e=>{const a=cur.filter(o=>o.etapa===e);return sum(a,"ad")+sum(a,"im")+sum(a,"mrr")+sum(a,"scc")}));
  $("funnel").innerHTML=`<div class="row head"><div>Etapa</div><div></div><div>Adesão + Impl.<br>RR · SCC</div></div>`+etapas.map(e=>{
    const a=cur.filter(o=>o.etapa===e), ad=sum(a,"ad"), im=sum(a,"im"), rr=sum(a,"mrr"), sc=sum(a,"scc");
    return `<div class="row" tabindex="0"><div>${esc(e)} <span style="color:var(--muted)">(${a.length})</span></div>
    <div class="track"><i class="c-ad" style="width:${ad/fmax*100}%"></i><i class="c-im" style="width:${im/fmax*100}%"></i><i class="c-mr" style="width:${rr/fmax*100}%"></i><i class="c-scc" style="width:${sc/fmax*100}%"></i></div>
    <div style="text-align:right">${brl(ad+im)}<br><small><span class="v-rr">RR ${brl(rr)}</span> · <span class="v-scc">SCC ${brl(sc)}</span></small></div>
    <div class="tip" role="tooltip"><b>${esc(e)}</b><span>${a.length} ${a.length===1?"oportunidade":"oportunidades"} · ${new Set(a.map(o=>o.acc)).size} contas</span>
      <div class="tl"><i class="c-ad"></i>Adesão (CDU)<em>${full(ad)}</em></div>
      <div class="tl"><i class="c-im"></i>Implantação (serviços)<em>${full(im)}</em></div>
      <div class="tl sep">Adesão + Implantação<em>${full(ad+im)}</em></div>
      <div class="tl"><i class="c-mr"></i>RR (mensalidade)<em>${full(rr)}</em></div>
      <div class="tl"><i class="c-scc"></i>SCC (serviços recorrentes)<em>${full(sc)}</em></div></div></div>`}).join("");

  // alertas
  const B=cur; const al=[], grp=a=>{const m={};a.forEach(o=>{(m[o.acc]=m[o.acc]||[]).push(o)});return Object.entries(m).map(([k,v])=>`${esc(title(k))} (${v.length})`).join(", ")};
  const a2=B.filter(o=>/^\s*[67]\./.test(o.etapa)&&o.p!=null&&o.p<=30); if(a2.length) al.push(["Negociação ou fechamento com probabilidade de 30% ou menos", grp(a2)]);
  const a3=B.filter(o=>o.p==null); if(a3.length) al.push(["Sem probabilidade preenchida", grp(a3)]);
  const a4=B.filter(o=>o.mi!=null&&o.mi<CUR); if(a4.length) al.push(["Data prevista já vencida", grp(a4)]);
  const a5=B.filter(o=>o.mi==null); if(a5.length) al.push(["Sem data prevista", grp(a5)]);
  const a6=B.filter(o=>o.ad+o.im+o.mrr+o.scc===0); if(a6.length) al.push(["Oportunidades sem valor", grp(a6)]);
  $("alerts").innerHTML= al.length? al.map(([t,d])=>`<li><b>${t}</b>${d}</li>`).join("") : `<li class="ok"><b>Nada a corrigir</b>Todas as oportunidades estão consistentes.</li>`;

  // executivos: um bloco por executivo com os três horizontes
  const blk=(k,a,ref,refLabel)=>{ const s=agg(a), oneTime=s.ad+s.im, refV=ref? sum(ref,"ad")+sum(ref,"im"):0, pct=refV? Math.round(oneTime/refV*100):null;
    return `<div class="xh ${k}"><div class="xt"><b>${HZ[k].name}</b><em>${s.n} opp · ${s.contas} contas</em></div>
      <div class="xv"><span><i class="dot d-ad"></i>Adesão</span><strong class="v-ad" title="${full(s.ad)}">${brl(s.ad)}</strong></div>
      <div class="xv"><span><i class="dot d-im"></i>Implantação</span><strong class="v-im" title="${full(s.im)}">${brl(s.im)}</strong></div>
      <div class="xv"><span><i class="dot d-rr"></i>RR mensal</span><strong class="v-rr" title="${full(s.mrr)}">${brl(s.mrr)}</strong></div>
      <div class="xv"><span><i class="dot d-scc"></i>SCC</span><strong class="v-scc" title="${full(s.scc)}">${brl(s.scc)}</strong></div>
      ${ref?`<div class="xshare">${pct==null?"—":pct+"%"} ${refLabel}<div class="xbar"><i style="width:${Math.min(100,pct||0)}%"></i></div></div>`:""}</div>`; };
  const ini=e=>{const p=execName(e).split(/\s+/).filter(w=>w.length>2);return ((p[0]||"?")[0]+((p[1]||"")[0]||"")).toUpperCase()};
  const xrow=(label,sub,a,cls,avatar)=>{ const pp=a.filter(o=>o.pp);
    return `<div class="xrow ${cls||""}"><div class="xwho"><span class="av">${avatar}</span><div><b>${esc(label)}</b><small>${sub}</small></div></div>
      ${blk("fc",a.filter(o=>o.fc),pp,"do pipeline (adesão + impl.)")}${blk("pp",pp,a,"do total (adesão + impl.)")}${blk("tt",a)}</div>`; };
  const rowsP=rows.filter(pOk);
  const ordem=[...execs].sort((x,y)=>{const f=e=>{const a=rowsP.filter(o=>o.exec===e);return [sum(a.filter(o=>o.fc),"ad")+sum(a.filter(o=>o.fc),"im"), sum(a.filter(o=>o.pp),"ad")+sum(a.filter(o=>o.pp),"im")]};const A=f(x),B=f(y);return B[0]-A[0]||B[1]-A[1]});
  $("execList").innerHTML=(execs.length>1? xrow(`Time (${execs.length} executivos)`, `${new Set(rowsP.map(o=>o.acc)).size} contas · ${rowsP.length} oportunidades`, rowsP, "team", "∑") : "")+
    ordem.map(e=>{const a=rowsP.filter(o=>o.exec===e);return xrow(nm(e), `${new Set(a.map(o=>o.acc)).size} contas · ${a.length} oportunidades`, a, "", ini(e))}).join("");

  // contas
  const q=state.q.toLowerCase(), byAcc={}; cur.forEach(o=>{(byAcc[o.acc]=byAcc[o.acc]||[]).push(o)});
  const peso=v=>sum(v,"ad")+sum(v,"im")+sum(v,"mrr")+sum(v,"scc");
  const accs=Object.entries(byAcc).filter(([k,v])=>!q||k.toLowerCase().includes(q)||v.some(o=>o.code.includes(q))).sort((a,b)=>peso(b[1])-peso(a[1]));
  $("accTbl").innerHTML=`<thead><tr><th>Conta</th><th class="n">Opp</th><th>Etapa mais avançada</th><th>Fechamento</th><th class="n h-ad">Adesão</th><th class="n h-im">Implantação</th><th class="n h-rr">RR</th><th class="n h-scc">SCC</th></tr></thead><tbody>`+
   (accs.length? accs.map(([k,v])=>{ const dmin=v.filter(o=>o.d).map(o=>o.d).sort((a,b)=>a-b)[0], open=state.open.has(k);
     let h=`<tr class="acc" data-a="${esc(k)}" tabindex="0" aria-expanded="${open}"><td><b style="font-weight:600">${esc(title(k))}</b></td><td class="n">${v.length}</td><td>${esc(topEtapa(v))}</td><td>${dmin?dmin.toLocaleDateString("pt-BR"):"—"}</td><td class="n v-ad">${brl(sum(v,"ad"))}</td><td class="n v-im">${brl(sum(v,"im"))}</td><td class="n v-rr">${brl(sum(v,"mrr"))}</td><td class="n v-scc">${brl(sum(v,"scc"))}</td></tr>`;
     if(open) h+=v.map(o=>`<tr class="det"><td>${esc(o.code)} · ${esc(o.desc)}<br>${esc(o.tipo)} · ${esc(nm(o.exec))}</td><td class="n"><span class="pill ${o.fc?"hi":""}">${o.p==null?"s/ prob.":o.p+"%"}</span></td><td>${esc(o.etapa)}</td><td>${o.d?o.d.toLocaleDateString("pt-BR"):"—"}</td><td class="n">${full(o.ad)}</td><td class="n">${full(o.im)}</td><td class="n">${full(o.mrr)}</td><td class="n">${full(o.scc)}</td></tr>`).join("");
     return h; }).join("") : `<tr><td colspan="8" style="color:var(--muted)">Nenhuma conta com estes filtros${q?" com essa busca":""}.</td></tr>`)+"</tbody>";
}


/* ---------- metas ---------- */
const VERT = [
  {k:"ad", m:"meta_adesao",      nome:"Adesão",      cls:"ad"},
  {k:"im", m:"meta_implantacao", nome:"Implantação", cls:"im"},
  {k:"mrr",m:"meta_rr",          nome:"RR mensal",   cls:"rr"},
  {k:"scc",m:"meta_scc",         nome:"SCC",         cls:"scc"}
];
const keyMi = mi => `${Math.floor(mi/12)}-${String(mi%12+1).padStart(2,"0")}`;
const CURK = keyMi(CUR);
function periodo(){
  const ks=[...state.meses].filter(k=>k!=="sem").sort();
  if(ks.length){ let label=rotuloDatas(arvoreDatas());
    if(/meses$/.test(label)) label=ks.map(k=>{ const [y,m]=k.split("-"); return MES[+m-1]+"/"+y.slice(2); }).join(", ");
    return {ks, label}; }
  return ks.length? {ks, label: rotuloDatas(arvoreDatas())} : {ks:[CURK], label:`${MESL[TODAY.getMonth()]}/${TODAY.getFullYear()} (mês atual)`};
}
const pct = (a,b) => b>0? a/b*100 : null;
const fmtPct = v => v==null? "—" : (v>=999? ">999" : Math.round(v))+"%";
const stat = v => v==null? ["sm","Sem meta"] : v>=100? ["ok","No caminho"] : v>=70? ["at","Atenção"] : ["rk","Em risco"];

function mesesPipe(P){
  const [y,m]=P.ks[0].split("-").map(Number); const ini=Math.max(CUR, y*12+m-1);
  return [0,1,2].map(i=>keyMi(ini+i));
}
const rotMes = k => { const [y,m]=k.split("-"); return MES[+m-1]+"/"+y.slice(2); };
function renderMetas(){ renderMetas0(); renderProj(); }
function renderMetas0(){
  const card=$("metasCard");
  if(!metasOn){ card.classList.remove("hidden"); $("metasDesc").textContent=""; $("metasBody").innerHTML=`<div class="empty" style="margin:0"><b>Metas ainda não ativadas no banco</b>Rode o script alteracao-metas.sql no Supabase.</div>`; return; }
  const P=periodo(), inP=o=>o.mi!=null && P.ks.includes(keyMi(o.mi));
  const P3=mesesPipe(P);
  // códigos visíveis
  let cods=new Set([...metas.map(m=>m.codigo_t.toUpperCase()), ...rows.map(o=>codigoT(o.exec)), ...ganhas.map(o=>codigoT(o.exec))].filter(Boolean));
  if(papel==="vendedor") cods=new Set(meuT?[meuT]:[]);
  else if(state.exec!=="__all"){ const c=codigoT(state.exec); cods=new Set(c?[c]:[]); }
  const nomeDe={}; rows.concat(ganhas).forEach(o=>{ const c=codigoT(o.exec); if(c && !nomeDe[c]) nomeDe[c]=nm(o.exec); });
  const calc = cs => { const r={};
    VERT.forEach(v=>{
      const meta = metas.filter(m=>cs.has(m.codigo_t.toUpperCase()) && P.ks.includes(String(m.mes).slice(0,7))).reduce((a,m)=>a+Number(m[v.m]||0),0);
      const real = ganhas.filter(o=>cs.has(codigoT(o.exec)) && inP(o)).reduce((a,o)=>a+o[v.k],0);
      const fc   = rows.filter(o=>cs.has(codigoT(o.exec)) && inP(o) && pOk0(o)).reduce((a,o)=>a+o[v.k],0);
      const meta3= metas.filter(m=>cs.has(m.codigo_t.toUpperCase()) && P3.includes(String(m.mes).slice(0,7))).reduce((a,m)=>a+Number(m[v.m]||0),0);
      const pipe3= rows.filter(o=>cs.has(codigoT(o.exec)) && o.pp).reduce((a,o)=>a+o[v.k],0);
      r[v.k]={meta,real,fc,at:pct(real,meta),proj:pct(real+fc,meta),gap:Math.max(0,meta-real-fc),cob: meta3>0? pipe3/meta3 : null};
    });
    const avg=f=>{ const xs=VERT.map(v=>r[v.k][f]).filter(x=>x!=null).map(x=>Math.min(100,x)); return xs.length? xs.reduce((a,b)=>a+b,0)/xs.length : null; };
    r.media = avg("at"); r.mproj = avg("proj");
    r.nG = ganhas.filter(o=>cs.has(codigoT(o.exec)) && inP(o)).length;
    return r; };
  const linha = (v,x) => { const [sc,st]=stat(x.proj), a=Math.min(100,x.at||0), f=Math.max(0,Math.min(100,(x.proj||0))-a);
    return `<div class="mv">
      <div class="mvl"><i class="dot d-${v.cls}"></i><div><b>${v.nome}</b><span class="mmeta">Meta ${x.meta? brl(x.meta) : "—"}</span></div></div>
      <div class="mvb"><div class="mbar" title="Realizado ${full(x.real)} · Previsto no período ${full(x.fc)} · Meta ${full(x.meta)}"><i class="mr ${sc}" style="width:${a}%"></i><i class="mf" style="width:${f}%"></i></div>
        <div class="mvt">${x.meta? `<b>${fmtPct(x.at)}</b> realizado · projeção <b>${fmtPct(x.proj)}</b> <span>(realizado + previsto)</span>` : `<span>sem meta no período${x.real+x.fc?` · realizado ${brl(x.real)} · previsto ${brl(x.fc)}`:""}</span>`}</div></div>
      <div class="mvn"><span class="mst ${sc}">${st}</span>
        <small>Meta <b>${brl(x.meta)}</b> · Realizado <b>${brl(x.real)}</b> · Previsto <b>${brl(x.fc)}</b></small>
        <small>${x.meta? (x.real+x.fc>=x.meta? `<span class="v-rr">Folga ${brl(x.real+x.fc-x.meta)}</span>` : `<span class="falta">Falta ${brl(x.meta-x.real-x.fc)}</span>`) : ""}${x.cob!=null?`${x.meta?" · ":""}pipe ${rotMes(P3[0]).split("/")[0]}–${rotMes(P3[2]).split("/")[0]} ${x.cob.toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1})}x`:""}</small></div>
    </div>`; };
  const bloco = (label,sub,avatar,cs,cls) => { const r=calc(cs);
    return {media:r.mproj, html:`<div class="mrow ${cls||""}"><div class="xwho"><span class="av">${avatar}</span><div><b>${esc(label)}</b><small>${sub}${r.nG?` · ${r.nG} ganha(s) no período`:""}</small>
      ${r.media!=null?`<span class="mmed ${stat(r.mproj)[0]}" title="Médias das vertentes com meta, cada uma limitada a 100%">Realizado ${fmtPct(r.media)} · Projeção ${fmtPct(r.mproj)}</span>`:""}</div></div>
      <div class="mvs">${VERT.map(v=>linha(v,r[v.k])).join("")}</div></div>`}; };
  const lista=[...cods].sort();
  const blocos=lista.map(c=>{ const n=nomeDe[c]||c; const ini=n.split(/\s+/).filter(w=>w.length>2).slice(0,2).map(w=>w[0]).join("").toUpperCase()||"T";
    return bloco(n, c, ini, new Set([c])); }).sort((a,b)=>(b.media??-1)-(a.media??-1));
  const time = (papel!=="vendedor" && lista.length>1) ? bloco(`Time (${lista.length} executivos)`, "soma de todos", "∑", new Set(lista), "team").html : "";
  card.classList.remove("hidden");
  $("metasDesc").innerHTML = `Período: <b>${esc(P.label)}</b> (${P.ks.length} ${P.ks.length>1?"meses":"mês"}; a meta é a soma dos meses) · muda pelo filtro de Data prevista. Na barra: cheio = realizado (ganhas, 100%); listrado = previsto (em andamento com data prevista no período${state.probs.size?", só as probabilidades filtradas":""}).`;
  $("metasBody").innerHTML = lista.length? time + blocos.map((b,i)=>b.html.replace('<span class="av">',`<span class="rank">${blocos.length>1?(i+1)+"º":""}</span><span class="av">`)).join("")
    : `<div class="empty" style="margin:0"><b>Nenhuma meta carregada</b>${papel==="vendedor"?"Sua meta ainda não foi cadastrada.":'Clique em "Carregar metas" para enviar a planilha de metas.'}</div>`;
}


/* ---------- projeções (velocímetros, rosca e pipe 3 meses) ---------- */
const COR={r:"#E24B4A",a:"#EF9F27",g:"#1D9E75",n:"#888780"};
function renderProj(){
  const card=$("projCard");
  if(!metasOn){ card.classList.add("hidden"); return; }
  card.classList.remove("hidden");
  const P=periodo(), inP=o=>o.mi!=null && P.ks.includes(keyMi(o.mi));
  let cods=new Set([...metas.map(m=>m.codigo_t.toUpperCase()), ...rows.map(o=>codigoT(o.exec)), ...ganhas.map(o=>codigoT(o.exec))].filter(Boolean));
  if(papel==="vendedor") cods=new Set(meuT?[meuT]:[]);
  else if(state.exec!=="__all"){ const c=codigoT(state.exec); cods=new Set(c?[c]:[]); }
  const meu=o=>cods.has(codigoT(o.exec));
  const metaDe=(v,ks)=>metas.filter(m=>cods.has(m.codigo_t.toUpperCase()) && ks.includes(String(m.mes).slice(0,7))).reduce((a,m)=>a+Number(m[v.m]||0),0);
  const prev=rows.filter(o=>meu(o) && inP(o) && pOk0(o));
  const quem = papel==="vendedor"? "" : state.exec!=="__all"? ` · ${esc(nm(state.exec))}` : (cods.size>1?` · time (${cods.size} executivos)`:"");
  $("projTit").innerHTML=`Projeção do período · ${esc(P.label)}${quem}`;
  $("projSub").textContent=`${P.ks.length} ${P.ks.length>1?"meses":"mês"} · meta = soma dos meses filtrados`;

  // velocímetros
  const MAX=150, pt=(p,r)=>{ const a=Math.PI*(1-Math.min(Math.max(p,0),MAX)/MAX); return [70+r*Math.cos(a),72-r*Math.sin(a)]; };
  const arc=(p0,p1,c)=>{ const [x0,y0]=pt(p0,52),[x1,y1]=pt(p1,52); return `<path d="M${x0.toFixed(1)} ${y0.toFixed(1)} A52 52 0 0 1 ${x1.toFixed(1)} ${y1.toFixed(1)}" fill="none" stroke="${c}" stroke-width="10"/>`; };
  $("gauges").innerHTML=VERT.map(v=>{
    const meta=metaDe(v,P.ks), real=ganhas.filter(o=>meu(o)&&inP(o)).reduce((a,o)=>a+o[v.k],0);
    const pv=prev.reduce((a,o)=>a+o[v.k],0), pond=prev.reduce((a,o)=>a+o[v.k]*((o.p||0)/100),0);
    if(!meta) return `<div class="gc sm"><h4>${v.nome}</h4><div class="gnone">Sem meta no período</div>
      <div class="kv"><span>Realizado</span><b>${brl(real)}</b><span>Previsto</span><b>${brl(pv)}</b></div></div>`;
    const proj=(real+pv)/meta*100, pp=(real+pond)/meta*100, [nx,ny]=pt(proj,44), [px,py]=pt(pp,44);
    const [sc]=stat(proj), dif=real+pv-meta;
    return `<div class="gc"><h4><i class="dot d-${v.cls}"></i> ${v.nome}</h4>
      <svg width="140" height="84" viewBox="0 0 140 84" role="img" aria-label="${v.nome}: projeção ${Math.round(proj)}%, ponderada ${Math.round(pp)}%">
      ${arc(0,70,COR.r)}${arc(70,100,COR.a)}${arc(100,150,COR.g)}
      <line x1="70" y1="72" x2="${px.toFixed(1)}" y2="${py.toFixed(1)}" stroke="var(--muted)" stroke-width="2" stroke-dasharray="3 3"/>
      <line x1="70" y1="72" x2="${nx.toFixed(1)}" y2="${ny.toFixed(1)}" stroke="var(--ink)" stroke-width="2.5"/><circle cx="70" cy="72" r="4" fill="var(--ink)"/></svg>
      <div class="gbig">${fmtPct(proj)}</div><div class="gsub">projeção · ponderada ${fmtPct(pp)}</div>
      <div class="kv"><span>Meta</span><b>${brl(meta)}</b><span>Realizado</span><b>${brl(real)}</b><span>Previsto</span><b>${brl(pv)}</b><span>Ponderado</span><b>${brl(pond)}</b></div>
      <span class="mst ${sc}" style="display:inline-block;margin-top:6px">${dif>=0?"Folga "+brl(dif):"Falta "+brl(-dif)}</span></div>`;
  }).join("");

  // rosca: origem do previsto por probabilidade
  const tot=o=>o.ad+o.im+o.mrr+o.scc, grupos={};
  prev.forEach(o=>{ const k=pKey(o); grupos[k]=(grupos[k]||0)+tot(o); });
  const total=Object.values(grupos).reduce((a,b)=>a+b,0);
  const corP=k=>k==="s"?COR.n : +k>=60?COR.g : +k>=30?COR.a : COR.r;
  const ks=Object.keys(grupos).sort((a,b)=>a==="s"?1:b==="s"?-1:b-a);
  if(!total){ $("donut").innerHTML=`<div class="gnone" style="padding:24px 0">Nada previsto no período${state.probs.size?" com as probabilidades filtradas":""}.</div>`; }
  else {
    const C=2*Math.PI*44; let off=0;
    const segs=ks.map(k=>{ const len=grupos[k]/total*C; const el=`<circle cx="60" cy="60" r="44" fill="none" stroke="${corP(k)}" stroke-width="16" stroke-dasharray="${len.toFixed(1)} ${C.toFixed(1)}" stroke-dashoffset="${(-off).toFixed(1)}" transform="rotate(-90 60 60)"/>`; off+=len; return el; }).join("");
    const madura=ks.filter(k=>k!=="s"&&+k>=60).reduce((a,k)=>a+grupos[k],0)/total;
    $("donut").innerHTML=`<div style="display:flex;align-items:center;gap:14px;flex-wrap:wrap">
      <svg width="120" height="120" viewBox="0 0 120 120" role="img" aria-label="Origem do previsto por probabilidade">
        <circle cx="60" cy="60" r="44" fill="none" stroke="var(--line)" stroke-width="16"/>${segs}
        <text x="60" y="57" text-anchor="middle" style="font-size:12px;font-weight:700;fill:var(--ink)">${brl(total)}</text>
        <text x="60" y="73" text-anchor="middle" style="font-size:11px;fill:var(--muted)">previsto</text></svg>
      <div class="dleg">${ks.map(k=>`<span><i style="background:${corP(k)}"></i>${k==="s"?"Sem probabilidade":k+"%"}: <b>${Math.round(grupos[k]/total*100)}%</b> · ${brl(grupos[k])}</span>`).join("")}
        <span class="mst ${madura>=.5?"ok":"at"}" style="align-self:flex-start;margin-top:4px">${madura>=.5?"Previsto maduro (60%+ é a maior parte)":"Mês depende de contas imaturas"}</span></div></div>`;
  }

  // pipe 3 meses
  const M3=mesesPipe(P), v=VERT.find(x=>x.k===state.pv)||VERT[0];
  const dados=M3.map(k=>({k, meta:metaDe(v,[k]), pipe:rows.filter(o=>meu(o) && o.mi!=null && keyMi(o.mi)===k).reduce((a,o)=>a+o[v.k],0)}));
  const sm=dados.reduce((a,d)=>a+d.meta,0), sp=dados.reduce((a,d)=>a+d.pipe,0), cob=sm? sp/sm : null;
  const cc=x=>x==null?"sm": x>=3?"ok": x>=2?"at":"rk";
  $("pipeTit").textContent=`Projeção do pipe · ${rotMes(M3[0])} a ${rotMes(M3[2])}`;
  $("pipeCob").className="mst "+cc(cob); $("pipeCob").textContent= cob==null? "Sem meta" : "Cobertura "+cob.toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1})+"x";
  $("pipeChips").innerHTML=VERT.map(x=>`<button type="button" class="chip ${x.k===v.k?"on":""}" data-pv="${x.k}">${x.nome}</button>`).join("");
  const mx=Math.max(1,...dados.map(d=>Math.max(d.meta,d.pipe))), h=100;
  $("pipeBars").innerHTML=dados.map((d,i)=>{ const x=70+i*170, hm=d.meta/mx*h, hp=d.pipe/mx*h, c=d.meta? d.pipe/d.meta : null;
    return `<rect x="${x}" y="${115-hm}" width="34" height="${Math.max(hm,1)}" rx="3" fill="var(--tt)"/><rect x="${x+40}" y="${115-hp}" width="34" height="${Math.max(hp,1)}" rx="3" fill="var(--c-${v.cls})"/>
      <text x="${x+17}" y="${Math.max(11,109-hm)}" text-anchor="middle" style="font-size:11px;fill:var(--muted)">${brl(d.meta).replace("R$ ","")}</text>
      <text x="${x+57}" y="${Math.max(11,109-hp)}" text-anchor="middle" style="font-size:11px;fill:var(--ink);font-weight:600">${brl(d.pipe).replace("R$ ","")}</text>
      <text x="${x+37}" y="133" text-anchor="middle" style="font-size:12px;fill:var(--muted)">${rotMes(d.k)}</text>
      <text x="${x+37}" y="148" text-anchor="middle" style="font-size:11.5px;font-weight:700;fill:${c==null?COR.n:c>=3?COR.g:c>=2?COR.a:COR.r}">${c==null?"sem meta":c.toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1})+"x"}</text>`; }).join("")+
    `<line x1="40" y1="115" x2="520" y2="115" stroke="var(--line)"/>`;
}

/* carga de metas */
const MREQ = ["Código T","Mês"], MVAL=["Meta Adesão","Meta Implantação","Meta RR","Meta SCC"];
const MALIAS = {"codigo t":"Código T","codigo":"Código T","cod t":"Código T","t":"Código T","vendedor":"Código T",
  "mes":"Mês","mes ano":"Mês","competencia":"Mês","periodo":"Mês",
  "meta adesao":"Meta Adesão","meta cdu":"Meta Adesão","meta cdu/adesao":"Meta Adesão","adesao":"Meta Adesão",
  "meta implantacao":"Meta Implantação","meta servicos":"Meta Implantação","meta servicos nao recorrentes":"Meta Implantação","implantacao":"Meta Implantação",
  "meta rr":"Meta RR","meta mensalidade":"Meta RR","rr":"Meta RR",
  "meta scc":"Meta SCC","meta servicos recorrentes":"Meta SCC","scc":"Meta SCC","nome":"Nome"};
const MCANON={}; Object.entries(MALIAS).forEach(([k,v])=>MCANON[normH(k)]=v);
const MABR={jan:1,fev:2,mar:3,abr:4,mai:5,jun:6,jul:7,ago:8,set:9,out:10,nov:11,dez:12};
function parseMes(v){
  if(v==null||v==="") return null;
  let y,m;
  if(v instanceof Date){ y=v.getFullYear(); m=v.getMonth()+1; }
  else if(typeof v==="number"){ const d=XLSX.SSF.parse_date_code(v); if(!d) return null; y=d.y; m=d.m; }
  else { const t=String(v).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
    let x;
    if((x=t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))){ y=+x[3]; m=+x[2]; }
    else if((x=t.match(/^(\d{1,2})[\/\-.](\d{4})$/))){ y=+x[2]; m=+x[1]; }
    else if((x=t.match(/^(\d{4})-(\d{1,2})/))){ y=+x[1]; m=+x[2]; }
    else if((x=t.match(/^([a-z]{3})[a-z]*[\/\-. ]+(\d{2,4})$/)) && MABR[x[1]]){ m=MABR[x[1]]; y=+x[2]; if(y<100) y+=2000; }
    else return null; }
  if(!(m>=1&&m<=12) || !(y>=2000&&y<=2100)) return null;
  return `${y}-${String(m).padStart(2,"0")}-01`;
}
let metasPend=null;
function abrirAjudaMetas(){ $("dlgMetas").showModal(); }
$("btnMetas").addEventListener("click", abrirAjudaMetas);
$("metasCancelar").addEventListener("click", ()=>$("dlgMetas").close());
$("metasEscolher").addEventListener("click", ()=>{ $("dlgMetas").close(); $("fileMetas").click(); });
$("metasModelo").addEventListener("click", ()=>{
  const cods=[...new Set(rows.concat(ganhas).map(o=>codigoT(o.exec)).filter(Boolean))].sort();
  const nomeDe={}; rows.concat(ganhas).forEach(o=>{ const c=codigoT(o.exec); if(c&&!nomeDe[c]) nomeDe[c]=nm(o.exec); });
  const linhas=[["Código T","Nome","Mês","Meta Adesão","Meta Implantação","Meta RR","Meta SCC"]];
  const lista=cods.length?cods:["T12345"];
  lista.forEach(c=>{ for(let mi=CUR; mi<=Math.floor(CUR/12)*12+11; mi++) linhas.push([c, nomeDe[c]||"", `${String(mi%12+1).padStart(2,"0")}/${Math.floor(mi/12)}`, 0,0,0,0]); });
  const ws=XLSX.utils.aoa_to_sheet(linhas); ws["!cols"]=[{wch:10},{wch:30},{wch:10},{wch:14},{wch:17},{wch:10},{wch:10}];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Metas"); XLSX.writeFile(wb, "modelo-metas.xlsx");
});
$("fileMetas").addEventListener("change", async e=>{
  const f=e.target.files[0]; e.target.value=""; if(!f) return;
  try{
    const wb=XLSX.read(await f.arrayBuffer(),{type:"array",cellDates:true});
    const raw=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:null}).map(r=>{ const o={}; for(const k in r){ const c=MCANON[normH(k)]||k; if(!(c in o)||o[c]==null) o[c]=r[k]; } return o; });
    if(!raw.length) throw new Error("A planilha de metas está vazia.");
    const cols=new Set(raw.flatMap(r=>Object.keys(r)));
    const miss=MREQ.filter(c=>!cols.has(c)); if(miss.length) throw new Error(`A planilha de metas não tem as colunas: ${miss.join(", ")}.`);
    if(!MVAL.some(c=>cols.has(c))) throw new Error("A planilha de metas não tem nenhuma coluna de meta (Meta Adesão, Meta Implantação, Meta RR ou Meta SCC).");
    const faltam=MVAL.filter(c=>!cols.has(c));
    const erros=[], mapa=new Map();
    raw.forEach((r,i)=>{
      const cv=r["Código T"], mes=parseMes(r["Mês"]);
      if((cv==null||cv==="") && (r["Mês"]==null||r["Mês"]==="")) return;
      let c=String(cv??"").trim().toUpperCase(); if(/^\d+$/.test(c)) c="T"+c;
      if(!/^T\d+$/.test(c)){ erros.push(`linha ${i+2}: código T inválido (${esc(cv??"vazio")})`); return; }
      if(!mes){ erros.push(`linha ${i+2}: mês inválido (${esc(r["Mês"]??"vazio")})`); return; }
      mapa.set(c+"|"+mes,{codigo_t:c,mes,meta_adesao:num(r["Meta Adesão"]),meta_implantacao:num(r["Meta Implantação"]),meta_rr:num(r["Meta RR"]),meta_scc:num(r["Meta SCC"])});
    });
    const linhas=[...mapa.values()];
    if(!linhas.length) throw new Error("Nenhuma linha válida na planilha de metas."+(erros.length?" "+erros.slice(0,3).join("; "):""));
    metasPend=linhas;
    const nomeDe={}; rows.concat(ganhas).forEach(o=>{ const c=codigoT(o.exec); if(c&&!nomeDe[c]) nomeDe[c]=nm(o.exec); });
    const porCod={}; linhas.forEach(l=>(porCod[l.codigo_t]=porCod[l.codigo_t]||[]).push(l));
    const mesTxt=k=>{ const [y,m]=k.split("-"); return MES[+m-1]+"/"+y; };
    $("metasConfBody").innerHTML=
      `<p style="font-size:13.5px;margin:0 0 10px">${esc(f.name)}: <b>${linhas.length}</b> meta(s) de <b>${Object.keys(porCod).length}</b> vendedor(es). As metas destes vendedores e meses serão substituídas; as demais não mudam.</p>`+
      (faltam.length?`<p style="font-size:13px;color:var(--muted)">Colunas ausentes (ficam zeradas): ${faltam.join(", ")}.</p>`:"")+
      (erros.length?`<p style="font-size:13px;color:var(--warn)">${erros.length} linha(s) ignorada(s): ${erros.slice(0,5).join("; ")}${erros.length>5?"…":""}</p>`:"")+
      `<div class="tbl"><table><thead><tr><th>Código T</th><th>Nome</th><th>Meses</th><th class="n h-ad">Adesão</th><th class="n h-im">Implantação</th><th class="n h-rr">RR</th><th class="n h-scc">SCC</th></tr></thead><tbody>`+
      Object.entries(porCod).sort().map(([c,ls])=>{ const ks=ls.map(l=>l.mes.slice(0,7)).sort(), t=k=>ls.reduce((a,l)=>a+l[k],0);
        return `<tr><td>${c}</td><td>${esc(nomeDe[c]||"—")}</td><td>${ks.length>1?mesTxt(ks[0])+" a "+mesTxt(ks[ks.length-1])+` (${ks.length})`:mesTxt(ks[0])}</td><td class="n">${brl(t("meta_adesao"))}</td><td class="n">${brl(t("meta_implantacao"))}</td><td class="n">${brl(t("meta_rr"))}</td><td class="n">${brl(t("meta_scc"))}</td></tr>`}).join("")+
      `</tbody></table></div><p style="font-size:12px;color:var(--muted)">Valores somados de todos os meses do arquivo.</p>`;
    $("dlgMetasConf").showModal();
  }catch(err){ say(err.message||"Não foi possível ler a planilha de metas.", false); }
});
$("metasConfCancelar").addEventListener("click", ()=>{ metasPend=null; $("dlgMetasConf").close(); });
$("metasConfirmar").addEventListener("click", async ()=>{
  if(!metasPend) return; const b=$("metasConfirmar"); b.disabled=true; b.textContent="Gravando…";
  const { data, error } = await db.rpc("registrar_metas",{ p_linhas: metasPend });
  b.disabled=false; b.textContent="Confirmar metas"; $("dlgMetasConf").close();
  if(error){ say(/Somente/i.test(error.message)? error.message : erroAmigavel(error), false); return; }
  say(`Metas gravadas: ${metasPend.length} linha(s).`, true); metasPend=null; await carregar();
});

/* ---------- interações ---------- */
$("exec").addEventListener("change",e=>{ state.exec=e.target.value; render(); });
$("q").addEventListener("input",e=>{ state.q=e.target.value; render(); $("q").focus(); });
document.addEventListener("change",e=>{
  const i=e.target.closest&&e.target.closest("input[data-node]"); if(!i) return;
  const ks=i.dataset.keys.split(",").filter(Boolean), cheio=ks.every(k=>state.meses.has(k));
  if(i.dataset.node==="all") { cheio? state.meses.clear() : ks.forEach(k=>state.meses.add(k)); }
  else cheio? ks.forEach(k=>state.meses.delete(k)) : ks.forEach(k=>state.meses.add(k));
  render();
});
document.addEventListener("click",e=>{
  const pv=e.target.closest("[data-pv]"); if(pv){ state.pv=pv.dataset.pv; renderProj(); return; }
  const dt=e.target.closest("[data-dt]"); if(dt){ state.dtOpen=!state.dtOpen; render(); return; }
  const ex=e.target.closest("[data-exp]"); if(ex){ const k=ex.dataset.exp; state.exp.has(k)?state.exp.delete(k):state.exp.add(k); render(); return; }
  if(state.dtOpen && !e.target.closest(".dpop")){ state.dtOpen=false; render(); }
  const h=e.target.closest("[data-h]"); if(h){ const k=h.dataset.h; state.hz = (state.hz===k || k==="tt") ? "tt" : k; render(); return; }
  const c=e.target.closest("[data-p]"); if(c){ const v=c.dataset.p;
    if(v==="__all") state.probs.clear(); else if(v==="__clear"){ state.probs.clear(); state.meses.clear(); state.hz="tt"; } else state.probs.has(v)? state.probs.delete(v) : state.probs.add(v);
    render(); return; }
  const t=e.target.closest(".tile"); if(t){ const k=t.dataset.acc; state.q=""; $("q").value=""; state.open.add(k); render();
    const r=[...document.querySelectorAll("tr.acc")].find(x=>x.dataset.a===k); if(r) r.scrollIntoView({behavior:"smooth",block:"center"}); return; }
  const a=e.target.closest("tr.acc"); if(a){ const k=a.dataset.a; state.open.has(k)?state.open.delete(k):state.open.add(k); render(); }
});
document.addEventListener("keydown",e=>{ const a=e.target.closest&&e.target.closest("tr.acc"); if(a&&(e.key==="Enter"||e.key===" ")){ e.preventDefault(); a.click(); } });

boot();
