/* Simulador de RV — página própria (somente dono). O salário nunca sai do navegador. */
const CFG = window.APP_CONFIG || {};
const db = window.supabase.createClient(CFG.SUPABASE_URL, CFG.SUPABASE_KEY);
const MES = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
const TODAY = new Date();
const CUR = TODAY.getFullYear()*12 + TODAY.getMonth();
const $ = id => document.getElementById(id);
const codigoT = e => { const m=String(e||"").match(/\(\s*(T\d+)\s*\)/i); return m? m[1].toUpperCase() : null; };
const MESL = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
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
const esc = s => String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
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
const keyMi = mi => `${Math.floor(mi/12)}-${String(mi%12+1).padStart(2,"0")}`;
let rows=[], ganhas=[], metas=[];

async function lerTudo(tabela){
  const out=[]; let from=0;
  while(true){ const { data, error } = await db.from(tabela).select("*").range(from, from+999); if(error) throw error; out.push(...data); if(data.length<1000) break; from+=1000; }
  return out;
}
function aviso(html){ $("rvPage").innerHTML=`<div class="empty" style="margin-top:20px">${html}</div>`; }

/* ---------- motor do Simulador RV (mesmas fórmulas da planilha) ---------- */
function calcRV(p, meses){
  // p: {fixo, mixFixo, gatilho, teto, ind:[{k,nome,peso,camp}]}
  // meses: [{meta:{nr,rr,sv}, real:{nr,rr,sv}}] x3
  const RV = p.fixo>0 ? p.fixo/p.mixFixo - p.fixo : 0;
  const fator=(real,meta,peso)=>{ if(!(meta>0)) return 0; const a=real/meta; return a<p.gatilho?0 : a>p.teto? p.teto*peso : a*peso; };
  const mensal=meses.map(m=>{ const por={}; p.ind.forEach(i=>por[i.k]=fator(m.real[i.k],m.meta[i.k],i.peso));
    const soma=p.ind.reduce((s,i)=>s+por[i.k],0); return {por, soma, valor: Math.min(soma*RV, RV)}; });
  const tri={}; p.ind.forEach(i=>{ const meta=meses.reduce((s,m)=>s+m.meta[i.k],0), real=meses.reduce((s,m)=>s+m.real[i.k],0);
    const f=fator(real,meta,i.peso), camp=(meta>0 && real/meta>=1)? f*(1+i.camp) : f;
    tri[i.k]={meta,real,at: meta>0? real/meta : null, f, camp, valor:f*RV*3, valorCamp:camp*RV*3}; });
  const somaTri=p.ind.reduce((s,i)=>s+tri[i.k].f,0), somaCamp=p.ind.reduce((s,i)=>s+tri[i.k].camp,0);
  const apuracaoTri = somaTri>p.teto ? 3*RV*p.teto : somaTri*3*RV;
  const apuracaoCamp = p.ind.reduce((s,i)=>s+tri[i.k].valorCamp,0);
  const pago = mensal.reduce((s,m)=>s+m.valor,0);
  return {RV, mensal, tri, somaTri, somaCamp, apuracaoTri, apuracaoCamp, pago, dif: apuracaoTri-pago, difCamp: apuracaoCamp-pago};
}
function calcSCC(s){ // s: {meta, real, gatilho, repIni, repFim, campanha, realTBC, acel}
  if(!(s.meta>0)) return {pct:0, valor:0, acel:0, total:0, at:null};
  const at=s.real/s.meta, passo=((s.repFim-s.repIni)/(1-s.gatilho))/100;
  const pct= at<s.gatilho? 0 : at>1? s.repFim : ((at-s.gatilho)*100)*passo + s.repIni;
  const valor=s.real*pct, acel= s.realTBC>=s.campanha && s.campanha>0 ? valor*s.acel : 0;
  return {at, pct, valor, acel, total: valor+acel};
}

/* ---------- simulador RV (somente dono, por enquanto) ---------- */
const RV_PADRAO={gatilho:0.4,teto:3,pesoNR:0.2,pesoRR:0.5,pesoSV:0.3,campNR:0,campRR:0.2,campSV:0.1,sccGat:0.4,sccIni:0.15,sccFim:0.5,sccCamp:130000,sccAcel:0.2};
let rvReg=(()=>{ try{ return {...RV_PADRAO, ...JSON.parse(localStorage.getItem("rv-regras")||"{}")}; }catch(e){ return {...RV_PADRAO}; } })();
let rvFixo=(()=>{ try{ return Number(localStorage.getItem("rv-fixo"))||0; }catch(e){ return 0; } })();   // só no navegador
let rvOver={}, rvTBC=null, rvSccMeta=null;
const brl2=v=>(v||0).toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
const pct1=v=>v==null?"—":(v*100).toLocaleString("pt-BR",{maximumFractionDigits:1})+"%";
const parseBR=v=>{ const t=String(v||"").replace(/[R$\s]/g,""); const n=t.includes(",")? Number(t.replace(/\./g,"").replace(",",".")) : Number(t); return Number.isFinite(n)?n:0; };
function rvInit(){
  const cods=[...new Set([...metas.map(m=>m.codigo_t.toUpperCase()), ...rows.concat(ganhas).map(o=>codigoT(o.exec))].filter(Boolean))].sort();
  const nomeDe={}; rows.concat(ganhas).forEach(o=>{ const c=codigoT(o.exec); if(c&&!nomeDe[c]) nomeDe[c]=nm(o.exec); });
  const atual=$("rvExec").value || (cods.includes("T30535")?"T30535":cods[0]);
  $("rvExec").innerHTML=cods.map(c=>`<option value="${c}" ${c===atual?"selected":""}>${esc(nomeDe[c]||c)} · ${c}</option>`).join("");
  const anos=[...new Set(metas.map(m=>+String(m.mes).slice(0,4)).concat([TODAY.getFullYear()]))].sort();
  const triAtual=`${TODAY.getFullYear()}-${Math.floor(TODAY.getMonth()/3)+1}`, sel=$("rvTri").value||triAtual;
  $("rvTri").innerHTML=anos.flatMap(a=>[1,2,3,4].map(q=>`<option value="${a}-${q}" ${`${a}-${q}`===sel?"selected":""}>${q}º trimestre de ${a}</option>`)).join("");
  $("rvFixo").value = rvFixo? rvFixo.toLocaleString("pt-BR",{minimumFractionDigits:2}) : "";
  try{ $("rvLembrar").checked = !!localStorage.getItem("rv-fixo"); }catch(e){}
  const campos=[["Gatilho (mínimo de atingimento)","gatilho",true],["Teto (máximo de atingimento)","teto",true],
    ["Peso Software NR (Adesão)","pesoNR",true],["Peso Software RR","pesoRR",true],["Peso Serviços NR (Implantação)","pesoSV",true],
    ["Campanha TRI Software NR","campNR",true],["Campanha TRI Software RR","campRR",true],["Campanha TRI Serviços","campSV",true],
    ["SCC · gatilho","sccGat",true],["SCC · repasse inicial","sccIni",true],["SCC · repasse final","sccFim",true],["SCC · meta campanha TBC (R$)","sccCamp",false],["SCC · acelerador TBC","sccAcel",true]];
  $("rvRegras").innerHTML='<h5>Plano de remuneração variável</h5>'+campos.slice(0,8).map(c=>rvCampo(c)).join("")+'<h5>Campanha SCC</h5>'+campos.slice(8).map(c=>rvCampo(c)).join("")+
    '<div style="grid-column:1/-1"><button type="button" id="rvPadrao">Voltar às regras padrão da planilha</button></div>';
}
function rvCampo([lab,k,isPct]){ const v=rvReg[k]; return `<label>${lab}<input data-rvreg="${k}" data-pct="${isPct?1:0}" value="${isPct? (v*100).toLocaleString("pt-BR",{maximumFractionDigits:2})+"%" : v.toLocaleString("pt-BR")}"></label>`; }
function rvDados(){
  const cod=$("rvExec").value, [ano,q]=$("rvTri").value.split("-").map(Number), cen=$("rvCen").value;
  const ks=[0,1,2].map(i=>`${ano}-${String((q-1)*3+i+1).padStart(2,"0")}`);
  const minP= cen==="real"? null : Number(cen);
  const meses=ks.map(k=>{
    const mt=metas.find(m=>m.codigo_t.toUpperCase()===cod && String(m.mes).slice(0,7)===k)||{};
    const g=ganhas.filter(o=>codigoT(o.exec)===cod && o.mi!=null && keyMi(o.mi)===k);
    const pv= minP==null? [] : rows.filter(o=>codigoT(o.exec)===cod && o.mi!=null && keyMi(o.mi)===k && (o.p||0)>=minP);
    const base={nr:sum(g,"ad")+sum(pv,"ad"), rr:sum(g,"mrr")+sum(pv,"mrr"), sv:sum(g,"im")+sum(pv,"im")};
    const real={}; ["nr","rr","sv"].forEach(x=>{ const o=rvOver[`${k}|${x}|${cod}|${cen}`]; real[x]= o!=null? o : base[x]; });
    return {k, meta:{nr:Number(mt.meta_adesao||0), rr:Number(mt.meta_rr||0), sv:Number(mt.meta_implantacao||0), scc:Number(mt.meta_scc||0)}, real, base,
      scc: sum(g,"scc")+sum(pv,"scc")};
  });
  return {cod, cen, meses};
}
function rvRender(){
  const {cod,cen,meses}=rvDados();
  const IND=[{k:"nr",nome:"Software NR (Adesão)",peso:rvReg.pesoNR,camp:rvReg.campNR},{k:"rr",nome:"Software RR (mensalidade)",peso:rvReg.pesoRR,camp:rvReg.campRR},{k:"sv",nome:"Serviços NR (Implantação)",peso:rvReg.pesoSV,camp:rvReg.campSV}];
  const r=calcRV({fixo:rvFixo, mixFixo:Number($("rvMix").value), gatilho:rvReg.gatilho, teto:rvReg.teto, ind:IND}, meses);
  const sMeta= rvSccMeta!=null? rvSccMeta : meses.reduce((a,m)=>a+m.meta.scc,0), sReal=meses.reduce((a,m)=>a+m.scc,0), tbc= rvTBC!=null? rvTBC : sReal;
  const sc=calcSCC({meta:sMeta, real:sReal, gatilho:rvReg.sccGat, repIni:rvReg.sccIni, repFim:rvReg.sccFim, campanha:rvReg.sccCamp, realTBC:tbc, acel:rvReg.sccAcel});
  const semMeta=meses.every(m=>!m.meta.nr&&!m.meta.rr&&!m.meta.sv);
  const semSal=!(rvFixo>0);
  const cenTxt=$("rvCen").selectedOptions[0].textContent;
  let h="";
  if(semSal) h+=`<div class="rvaviso">Digite o seu <b>salário fixo mensal</b> acima para ver os valores em reais. Os percentuais de atingimento já aparecem abaixo.</div>`;
  if(semMeta) h+=`<div class="rvaviso">Não há metas cadastradas para ${cod} neste trimestre. Carregue as metas em "Carregar metas".</div>`;
  h+=`<div class="rvcards">
    <div class="rvcard"><span>Referência variável mensal</span><b>${semSal?"—":brl2(r.RV)}</b></div>
    <div class="rvcard"><span>Pago mês a mês (soma do tri)</span><b>${semSal?"—":brl2(r.pago)}</b></div>
    <div class="rvcard"><span>Ajuste no fechamento do tri</span><b class="${r.difCamp<0?"rvneg":"rvpos"}">${semSal?"—":brl2(r.difCamp)}</b></div>
    <div class="rvcard"><span>Campanha SCC</span><b>${semSal?"—":brl2(sc.total)}</b></div>
    <div class="rvcard dest"><span>Variável total no tri (com campanhas)</span><b>${semSal?pct1(r.somaCamp)+" da RV":brl2(r.apuracaoCamp+sc.total)}</b></div>
  </div>`;
  const nomeMes=k=>MESL[+k.slice(5,7)-1]+"/"+k.slice(0,4);
  h+=`<div class="rvmeses">`+meses.map((m,i)=>{ const mr=r.mensal[i];
    return `<div class="rvtab"><h4>${nomeMes(m.k)}</h4><table><thead><tr><th>Indicador</th><th class="n">Meta</th><th class="n">Realizado</th><th class="n">Ating.</th><th class="n">Fator</th></tr></thead><tbody>`+
      IND.map(x=>{ const ov=rvOver[`${m.k}|${x.k}|${cod}|${cen}`]!=null, at=m.meta[x.k]>0? m.real[x.k]/m.meta[x.k] : null;
        return `<tr><td>${x.nome.split(" (")[0]}</td><td class="n">${brl(m.meta[x.k])}</td>
          <td class="n"><input class="${ov?"ov":""}" data-rvreal="${m.k}|${x.k}" value="${m.real[x.k].toLocaleString("pt-BR",{maximumFractionDigits:2})}" title="Valor do painel: ${brl2(m.base[x.k])}. Edite para simular."></td>
          <td class="n ${at!=null&&at<rvReg.gatilho?"rvneg":""}">${pct1(at)}</td><td class="n">${pct1(mr.por[x.k])}</td></tr>`; }).join("")+
      `<tr class="rvtot"><td colspan="4">Apuração do mês</td><td class="n">${pct1(mr.soma)}</td></tr>
       <tr class="rvtot"><td colspan="4">Valor a ser pago (até 100% da RV)</td><td class="n">${semSal?"—":brl2(mr.valor)}</td></tr></tbody></table></div>`; }).join("")+`</div>`;
  h+=`<div class="rvsec"><h4>Balanço do trimestre</h4><div class="tbl"><table><thead><tr><th>Indicador</th><th class="n">Peso</th><th class="n">Meta tri</th><th class="n">Realizado tri</th><th class="n">Ating.</th><th class="n">Fator</th><th class="n">Valor apurado</th><th class="n">Fator c/ campanha</th><th class="n">Valor c/ campanha</th></tr></thead><tbody>`+
    IND.map(x=>{ const t=r.tri[x.k]; return `<tr><td>${x.nome}</td><td class="n">${pct1(x.peso)}</td><td class="n">${brl(t.meta)}</td><td class="n">${brl(t.real)}</td><td class="n ${t.at!=null&&t.at<rvReg.gatilho?"rvneg":""}">${pct1(t.at)}</td><td class="n">${pct1(t.f)}</td><td class="n">${semSal?"—":brl2(t.valor)}</td><td class="n">${pct1(t.camp)}${t.camp>t.f?' <span class="rvpos">▲</span>':""}</td><td class="n">${semSal?"—":brl2(t.valorCamp)}</td></tr>`; }).join("")+
    `<tr class="rvtot"><td colspan="5">Atingimento do tri</td><td class="n">${pct1(r.somaTri)}</td><td class="n">${semSal?"—":brl2(r.apuracaoTri)}</td><td class="n">${pct1(r.somaCamp)}</td><td class="n">${semSal?"—":brl2(r.apuracaoCamp)}</td></tr>
     <tr><td colspan="6">Já pago mês a mês</td><td class="n">${semSal?"—":brl2(r.pago)}</td><td></td><td class="n">${semSal?"—":brl2(r.pago)}</td></tr>
     <tr class="rvtot"><td colspan="6">Diferença no fechamento do tri</td><td class="n ${r.dif<0?"rvneg":"rvpos"}">${semSal?"—":brl2(r.dif)}</td><td></td><td class="n ${r.difCamp<0?"rvneg":"rvpos"}">${semSal?"—":brl2(r.difCamp)}</td></tr>
    </tbody></table></div>
    <p style="font-size:12px;color:var(--muted);margin:6px 0 0">Cenário: ${esc(cenTxt)}. Realizado = ganhas (100%) no mês da data prevista${cen==="real"?"":" + oportunidades em andamento do cenário"}. Campos em azul foram editados por você. Abaixo do gatilho (${pct1(rvReg.gatilho)}), o indicador não paga; o teto é ${pct1(rvReg.teto)}.</p></div>`;
  h+=`<div class="rvsec"><h4>Campanha SCC do trimestre</h4><div class="tbl"><table><thead><tr><th class="n">Meta tri</th><th class="n">Realizado tri</th><th class="n">Ating.</th><th class="n">% de repasse</th><th class="n">Valor apurado</th><th class="n">Realizado TBC (campanha ${brl(rvReg.sccCamp)})</th><th class="n">Acelerador</th><th class="n">Total</th></tr></thead><tbody>
    <tr><td class="n"><input data-rvscc="meta" value="${sMeta.toLocaleString("pt-BR")}" style="width:110px;text-align:right"></td><td class="n">${brl2(sReal)}</td><td class="n">${pct1(sc.at)}</td><td class="n">${pct1(sc.pct)}</td><td class="n">${brl2(sc.valor)}</td>
    <td class="n"><input data-rvscc="tbc" value="${tbc.toLocaleString("pt-BR")}" style="width:120px;text-align:right"></td><td class="n">${brl2(sc.acel)}</td><td class="n"><b>${brl2(sc.total)}</b></td></tr></tbody></table></div>
    <p style="font-size:12px;color:var(--muted);margin:6px 0 0">O SCC paga um percentual sobre o realizado: ${pct1(rvReg.sccIni)} ao atingir o gatilho de ${pct1(rvReg.sccGat)}, subindo até ${pct1(rvReg.sccFim)} em 100%. O acelerador de ${pct1(rvReg.sccAcel)} vale quando o realizado TBC atinge a meta da campanha.</p></div>`;
  if(!semSal) h+=`<div class="rvsec"><h4>Remuneração estimada no trimestre</h4><p style="font-size:13.5px;margin:0">Fixo ${brl2(rvFixo*3)} + variável ${brl2(r.apuracaoCamp)} + SCC ${brl2(sc.total)} = <b>${brl2(rvFixo*3+r.apuracaoCamp+sc.total)}</b> (valores brutos, antes de impostos).</p></div>`;
  $("rvOut").innerHTML=h;
}
$("rvOlho").addEventListener("click", ()=>{ const i=$("rvFixo"); i.type= i.type==="password"? "text" : "password"; });
$("rvRestaurar").addEventListener("click", ()=>{ rvOver={}; rvTBC=null; rvSccMeta=null; rvRender(); });
["rvExec","rvTri","rvCen","rvMix"].forEach(id=>$(id).addEventListener("change", ()=>{ rvTBC=null; rvSccMeta=null; rvRender(); }));
$("rvFixo").addEventListener("change", ()=>{ rvFixo=parseBR($("rvFixo").value); if($("rvLembrar").checked){ try{ localStorage.setItem("rv-fixo", String(rvFixo)); }catch(e){} } rvRender(); });
$("rvLembrar").addEventListener("change", e=>{ try{ e.target.checked? localStorage.setItem("rv-fixo", String(rvFixo)) : localStorage.removeItem("rv-fixo"); }catch(err){} });
document.addEventListener("change", e=>{
  const t=e.target;
  if(t.dataset.rvreal){ const {cod,cen}={cod:$("rvExec").value,cen:$("rvCen").value}; rvOver[`${t.dataset.rvreal}|${cod}|${cen}`]=parseBR(t.value); rvRender(); }
  else if(t.dataset.rvscc){ const v=parseBR(t.value); if(t.dataset.rvscc==="tbc") rvTBC=v; else rvSccMeta=v; rvRender(); }
  else if(t.dataset.rvreg){ let v=parseBR(String(t.value).replace("%","")); if(t.dataset.pct==="1") v=v/100; rvReg[t.dataset.rvreg]=v; try{ localStorage.setItem("rv-regras", JSON.stringify(rvReg)); }catch(err){} rvRender(); }
});
document.addEventListener("click", e=>{ if(e.target.id==="rvPadrao"){ rvReg={...RV_PADRAO}; try{ localStorage.removeItem("rv-regras"); }catch(err){} rvInit(); rvRender(); } });


(async function iniciar(){
  if(CFG.LOGO){ const i=$("logo"); if(i){ i.src=CFG.LOGO; i.classList.add("on"); } }
  const { data:{ session } } = await db.auth.getSession();
  if(!session){ aviso('<b>Você não está conectado</b>Entre no <a href="./">painel</a> e clique de novo em "Simulador RV".'); return; }
  const { data:perfil } = await db.rpc("meu_perfil");
  if(!perfil || perfil.papel!=="dono"){ aviso('<b>Acesso restrito</b>Por enquanto, o simulador está disponível apenas para o dono do painel.'); return; }
  try{
    const [a,g,m]=await Promise.all([lerTudo("oportunidades_atuais"), lerTudo("oportunidades_ganhas"), lerTudo("metas")]);
    rows=a.map(toView).filter(o=>o.p!==100);
    ganhas=g.map(r=>{ const o=toView(r); if(o.mi==null && r.carga_em){ const d=new Date(r.carga_em); o.mi=d.getFullYear()*12+d.getMonth(); } return o; });
    metas=m;
  }catch(e){ aviso('<b>Não foi possível carregar os dados</b>'+esc(e.message||String(e))); return; }
  $("rvPage").classList.remove("hidden"); $("rvCarregando").classList.add("hidden");
  rvInit(); rvRender();
})();
