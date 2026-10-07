/* Simulador de RV — página própria. Acesso: dono ou usuário liberado (acesso_rv). Quem não é dono simula só a própria carteira.
   O salário nunca sai do navegador; se "lembrar" estiver marcado, fica guardado por login e é apagado ao sair do painel. */
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
function aviso(html){ $("rvPage").innerHTML=`<div class="empty" style="margin-top:20px">${html}</div>`; $("rvPage").classList.remove("hidden"); $("rvCarregando").classList.add("hidden"); }

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

/* ---------- formato da planilha e calendário de pagamento ---------- */
const pct2=v=>v==null?"—":(v*100).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})+"%";   // indicadores: 60,00%
const pct0=v=>v==null?"—":Math.round(v*100).toLocaleString("pt-BR")+"%";                                          // totais: 300%
function pascoa(a){ const b=a%19,c=Math.floor(a/100),d=a%100,e=Math.floor(c/4),f=c%4,g=Math.floor((c+8)/25),h=Math.floor((c-g+1)/3),
  i=(19*b+c-e-h+15)%30,k=Math.floor(d/4),l=d%4,m=(32+2*f+2*k-i-l)%7,n=Math.floor((b+11*i+22*m)/451),mes=Math.floor((i+m-7*n+114)/31),dia=((i+m-7*n+114)%31)+1;
  return new Date(a,mes-1,dia); }
function feriados(a){   // feriados nacionais + carnaval, sexta-feira santa e Corpus Christi (bancos fechados)
  const f=new Set(["01-01","04-21","05-01","09-07","10-12","11-02","11-15","11-20","12-25"].map(x=>`${a}-${x}`));
  const p=pascoa(a); [-48,-47,-2,60].forEach(d=>{ const x=new Date(p); x.setDate(x.getDate()+d); f.add(iso(x)); });
  return f; }
function quintoDiaUtil(a, m){   // m = 1..12; dias úteis = segunda a sexta, sem feriado
  const f=feriados(a); let n=0;
  for(let d=1; d<=31; d++){ const x=new Date(a,m-1,d); if(x.getMonth()!==m-1) break; const w=x.getDay();
    if(w>0 && w<6 && !f.has(iso(x)) && ++n===5) return x; }
  return null; }
const pagamentoDe = k => { const a=+k.slice(0,4), m=+k.slice(5,7)+2; return quintoDiaUtil(a+Math.floor((m-1)/12), ((m-1)%12)+1); };   // competência + 2 meses
const dataBR = d => d? d.toLocaleDateString("pt-BR",{weekday:"short",day:"2-digit",month:"2-digit",year:"numeric"}) : "—";

/* ---------- simulador RV ---------- */
let SOU_DONO=false, MEU_COD=null, CHAVE_FIXO="rv-fixo";   // a chave do salário é por login (definida ao entrar)
const lerFixo = () => { try{ return Number(localStorage.getItem(CHAVE_FIXO))||0; }catch(e){ return 0; } };
const MIX_FIXO=0.6;   // plano 60% fixo / 40% variável (padrão da planilha)
const RV_PADRAO={gatilho:0.4,teto:3,pesoNR:0.2,pesoRR:0.5,pesoSV:0.3,campNR:0,campRR:0.2,campSV:0.1,sccGat:0.4,sccIni:0.15,sccFim:0.5,sccCamp:130000,sccAcel:0.2};
let rvReg=(()=>{ try{ return {...RV_PADRAO, ...JSON.parse(localStorage.getItem("rv-regras")||"{}")}; }catch(e){ return {...RV_PADRAO}; } })();
let rvFixo=0;   // só no navegador
let rvOver={}, rvTBC=null, rvSccMeta=null;
const brl2=v=>(v||0).toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
const pct1=v=>v==null?"—":(v*100).toLocaleString("pt-BR",{maximumFractionDigits:1})+"%";
const parseBR=v=>{ const t=String(v||"").replace(/[R$\s]/g,""); const n=t.includes(",")? Number(t.replace(/\./g,"").replace(",",".")) : Number(t); return Number.isFinite(n)?n:0; };
function rvInit(){
  const cods = SOU_DONO? [...new Set([...metas.map(m=>m.codigo_t.toUpperCase()), ...rows.concat(ganhas).map(o=>codigoT(o.exec))].filter(Boolean))].sort() : [MEU_COD];
  const nomeDe={}; rows.concat(ganhas).forEach(o=>{ const c=codigoT(o.exec); if(c&&!nomeDe[c]) nomeDe[c]=nm(o.exec); });
  const atual=$("rvExec").value || (cods.includes("T30535")?"T30535":cods[0]);
  $("rvExec").innerHTML=cods.map(c=>`<option value="${c}" ${c===atual?"selected":""}>${esc(nomeDe[c]||c)} · ${c}</option>`).join("");
  $("rvExec").disabled = !SOU_DONO;
  const anos=[...new Set(metas.map(m=>+String(m.mes).slice(0,4)).concat([TODAY.getFullYear()]))].sort();
  const triAtual=`${TODAY.getFullYear()}-${Math.floor(TODAY.getMonth()/3)+1}`, sel=$("rvTri").value||triAtual;
  $("rvTri").innerHTML=anos.flatMap(a=>[1,2,3,4].map(q=>`<option value="${a}-${q}" ${`${a}-${q}`===sel?"selected":""}>${q}º trimestre de ${a}</option>`)).join("");
  $("rvFixo").value = rvFixo? rvFixo.toLocaleString("pt-BR",{minimumFractionDigits:2}) : "";
  try{ $("rvLembrar").checked = !!localStorage.getItem(CHAVE_FIXO); }catch(e){}
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
  const r=calcRV({fixo:rvFixo, mixFixo:MIX_FIXO, gatilho:rvReg.gatilho, teto:rvReg.teto, ind:IND}, meses);
  const sMeta= rvSccMeta!=null? rvSccMeta : meses.reduce((a,m)=>a+m.meta.scc,0), sReal=meses.reduce((a,m)=>a+m.scc,0), tbc= rvTBC!=null? rvTBC : sReal;
  const sc=calcSCC({meta:sMeta, real:sReal, gatilho:rvReg.sccGat, repIni:rvReg.sccIni, repFim:rvReg.sccFim, campanha:rvReg.sccCamp, realTBC:tbc, acel:rvReg.sccAcel});
  const semMeta=meses.every(m=>!m.meta.nr&&!m.meta.rr&&!m.meta.sv);
  const semSal=!(rvFixo>0);
  const cenTxt=$("rvCen").selectedOptions[0].textContent;
  let h="";
  if(semSal) h+=`<div class="rvaviso">Digite o seu <b>salário fixo mensal</b> acima para ver os valores em reais. Os percentuais de atingimento já aparecem abaixo.</div>`;
  if(semMeta) h+=`<div class="rvaviso">Não há metas cadastradas para ${cod} neste trimestre. Carregue as metas em "Carregar metas".</div>`;
  const ult=meses[meses.length-1].k, nomeCurto=k=>MESL[+k.slice(5,7)-1].toLowerCase();
  h+=`<div class="rvcards">
    <div class="rvcard"><span>RV mensal (100% da meta)</span><b>${semSal?"—":brl2(r.RV)}</b></div>
    ${meses.map((m,i)=>`<div class="rvcard"><span>RV de ${nomeCurto(m.k)} · recebe ${pagamentoDe(m.k)?pagamentoDe(m.k).toLocaleDateString("pt-BR"):"—"}</span><b>${semSal?pct0(Math.min(r.mensal[i].soma,1)):brl2(r.mensal[i].valor)}</b></div>`).join("")}
    <div class="rvcard dest"><span>Fechamento do tri · recebe ${pagamentoDe(ult)?pagamentoDe(ult).toLocaleDateString("pt-BR"):"—"}</span><b class="${r.difCamp<0?"rvneg":""}">${semSal?"—":brl2(r.difCamp+sc.total)}</b></div>
  </div>`;
  const nomeMes=k=>MESL[+k.slice(5,7)-1]+"/"+k.slice(0,4);
  h+=`<div class="rvmeses">`+meses.map((m,i)=>{ const mr=r.mensal[i];
    return `<div class="rvtab"><h4>${nomeMes(m.k)}</h4><table><thead><tr><th>Indicador</th><th class="n">Meta</th><th class="n">Realizado</th><th class="n" title="Quanto do valor da meta foi atingido">% meta</th><th class="n" title="Percentual já ponderado pelo peso do indicador (é o que vira valor)">% peso</th><th class="n">Valor</th></tr></thead><tbody>`+
      IND.map(x=>{ const ov=rvOver[`${m.k}|${x.k}|${cod}|${cen}`]!=null, at=m.meta[x.k]>0? m.real[x.k]/m.meta[x.k] : null, abaixo=at!=null&&at<rvReg.gatilho;
        return `<tr><td>${x.nome.split(" (")[0]}</td><td class="n">${brl2(m.meta[x.k])}</td>
          <td class="n"><input class="${ov?"ov":""}" data-rvreal="${m.k}|${x.k}" value="${m.real[x.k].toLocaleString("pt-BR",{maximumFractionDigits:2})}" title="Valor do painel: ${brl2(m.base[x.k])}. Edite para simular."></td>
          <td class="n ${abaixo?"rvneg":""}" title="${abaixo?"Abaixo do gatilho de "+pct1(rvReg.gatilho)+": não paga":at!=null&&at>rvReg.teto?"Acima do teto de "+pct1(rvReg.teto)+": o % peso usa o teto":""}">${pct2(at)}</td>
          <td class="n" title="Peso ${pct1(x.peso)}">${pct2(mr.por[x.k])}</td><td class="n">${semSal?"—":brl2(mr.por[x.k]*r.RV)}</td></tr>`; }).join("")+
      `<tr class="rvtot"><td colspan="4">Apuração mensal</td><td class="n">${pct0(mr.soma)}</td><td class="n">${semSal?"—":brl2(mr.soma*r.RV)}</td></tr>
       <tr class="rvtot rvpagar"><td colspan="4">Valor a ser pago <small>· recebe ${dataBR(pagamentoDe(m.k))}</small></td><td></td><td class="n">${semSal?"—":brl2(mr.valor)}</td></tr></tbody></table></div>`; }).join("")+`</div>`;
  h+=`<div class="rvsec"><h4>Balanço do trimestre</h4><div class="tbl"><table><thead><tr><th>Indicador</th><th class="n">Meta</th><th class="n">Realizado</th><th class="n" title="Quanto do valor da meta foi atingido">% meta</th><th class="n" title="Percentual já ponderado pelo peso do indicador">% peso</th><th class="n">Valor apurado</th><th class="n">% Campanha</th><th class="n">Valor c/ campanha</th></tr></thead><tbody>`+
    IND.map(x=>{ const t=r.tri[x.k], abaixo=t.at!=null&&t.at<rvReg.gatilho; return `<tr><td>${x.nome}</td><td class="n">${brl2(t.meta)}</td><td class="n">${brl2(t.real)}</td>
      <td class="n ${abaixo?"rvneg":""}">${pct2(t.at)}</td><td class="n" title="Peso ${pct1(x.peso)}">${pct2(t.f)}</td><td class="n">${semSal?"—":brl2(t.valor)}</td>
      <td class="n">${pct2(t.camp)}${t.camp>t.f?' <span class="rvpos">▲</span>':""}</td><td class="n">${semSal?"—":brl2(t.valorCamp)}</td></tr>`; }).join("")+
    `<tr class="rvtot"><td colspan="4">Atingimento TRI</td><td class="n">${pct0(r.somaTri)}</td><td></td><td class="n">${pct0(r.somaCamp)}</td><td></td></tr>
     <tr><td colspan="4">Apuração TRI</td><td></td><td class="n">${semSal?"—":brl2(r.apuracaoTri)}</td><td></td><td class="n">${semSal?"—":brl2(r.apuracaoCamp)}</td></tr>
     <tr><td colspan="4">Valor pago (mês a mês)</td><td></td><td class="n">${semSal?"—":brl2(r.pago)}</td><td></td><td class="n">${semSal?"—":brl2(r.pago)}</td></tr>
     <tr class="rvtot"><td colspan="4">Diferença</td><td></td><td class="n ${r.dif<0?"rvneg":"rvpos"}">${semSal?"—":brl2(r.dif)}</td><td></td><td class="n ${r.difCamp<0?"rvneg":"rvpos"}">${semSal?"—":brl2(r.difCamp)}</td></tr>
    </tbody></table></div>
    <p style="font-size:12px;color:var(--muted);margin:6px 0 0">Cenário: ${esc(cenTxt)}. Realizado = ganhas (100%) no mês da data prevista${cen==="real"?"":" + oportunidades em andamento do cenário"}. Campos em azul foram editados por você. "% meta" é quanto da meta do indicador foi atingido; "% peso" é esse atingimento multiplicado pelo peso do indicador, e é ele que vira valor. Abaixo do gatilho (${pct1(rvReg.gatilho)} da meta) o indicador não paga; acima do teto (${pct1(rvReg.teto)}) o % peso para de subir.</p></div>`;
  h+=`<div class="rvsec"><h4>Campanha SCC do trimestre</h4><div class="tbl"><table><thead><tr><th class="n">Meta tri</th><th class="n">Realizado tri</th><th class="n">Ating.</th><th class="n">% de repasse</th><th class="n">Valor apurado</th><th class="n">Realizado TBC (campanha ${brl(rvReg.sccCamp)})</th><th class="n">Acelerador</th><th class="n">Total</th></tr></thead><tbody>
    <tr><td class="n"><input data-rvscc="meta" value="${sMeta.toLocaleString("pt-BR")}" style="width:110px;text-align:right"></td><td class="n">${brl2(sReal)}</td><td class="n">${pct1(sc.at)}</td><td class="n">${pct1(sc.pct)}</td><td class="n">${brl2(sc.valor)}</td>
    <td class="n"><input data-rvscc="tbc" value="${tbc.toLocaleString("pt-BR")}" style="width:120px;text-align:right"></td><td class="n">${brl2(sc.acel)}</td><td class="n"><b>${brl2(sc.total)}</b></td></tr></tbody></table></div>
    <p style="font-size:12px;color:var(--muted);margin:6px 0 0">O SCC paga um percentual sobre o realizado: ${pct1(rvReg.sccIni)} ao atingir o gatilho de ${pct1(rvReg.sccGat)}, subindo até ${pct1(rvReg.sccFim)} em 100%. O acelerador de ${pct1(rvReg.sccAcel)} vale quando o realizado TBC atinge a meta da campanha.</p></div>`;
  if(!semSal){ const dUlt=pagamentoDe(ult);
    const lin=meses.map((m,i)=>[`RV de ${nomeMes(m.k)}`, pagamentoDe(m.k), r.mensal[i].valor]);
    lin.push(["Diferença do fechamento do tri (com campanhas)", dUlt, r.difCamp]); if(sc.total) lin.push(["Campanha SCC do tri", dUlt, sc.total]);
    const porData={}; lin.forEach(([,d,v])=>{ const k=iso(d); porData[k]=(porData[k]||0)+v; });
    h+=`<div class="rvsec"><h4>Calendário de recebimento</h4><div class="tbl"><table><thead><tr><th>O que é</th><th>Recebe em (5º dia útil)</th><th class="n">Valor</th></tr></thead><tbody>`+
      lin.map(([t,d,v])=>`<tr><td>${t}</td><td>${dataBR(d)}</td><td class="n ${v<0?"rvneg":""}">${brl2(v)}</td></tr>`).join("")+
      `<tr class="rvtot"><td colspan="2">Total variável do trimestre</td><td class="n">${brl2(lin.reduce((a,x)=>a+x[2],0))}</td></tr></tbody></table></div>
      <p style="font-size:12px;color:var(--muted);margin:6px 0 0">Cada mês é pago no 5º dia útil do segundo mês seguinte (ex.: setembro → novembro). Dias úteis de segunda a sexta, sem feriados nacionais. Em ${dataBR(dUlt)} entram juntos ${brl2(porData[iso(dUlt)])}.</p></div>`; }
  if(!semSal) h+=`<div class="rvsec"><h4>Remuneração estimada no trimestre</h4><p style="font-size:13.5px;margin:0">Fixo ${brl2(rvFixo*3)} + variável ${brl2(r.apuracaoCamp)} + SCC ${brl2(sc.total)} = <b>${brl2(rvFixo*3+r.apuracaoCamp+sc.total)}</b> (valores brutos, antes de impostos).</p></div>`;
  $("rvOut").innerHTML=h;
}
$("rvOlho").addEventListener("click", ()=>{ const i=$("rvFixo"); i.type= i.type==="password"? "text" : "password"; });
$("rvRestaurar").addEventListener("click", ()=>{ rvOver={}; rvTBC=null; rvSccMeta=null; rvRender(); });
["rvExec","rvTri","rvCen"].forEach(id=>$(id).addEventListener("change", ()=>{ rvTBC=null; rvSccMeta=null; rvRender(); }));
$("rvFixo").addEventListener("change", ()=>{ rvFixo=parseBR($("rvFixo").value); if($("rvLembrar").checked){ try{ localStorage.setItem(CHAVE_FIXO, String(rvFixo)); }catch(e){} } rvRender(); });
$("rvLembrar").addEventListener("change", e=>{ try{ e.target.checked? localStorage.setItem(CHAVE_FIXO, String(rvFixo)) : localStorage.removeItem(CHAVE_FIXO); }catch(err){} });
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
  SOU_DONO = !!perfil && perfil.papel==="dono";
  if(!SOU_DONO){
    const { data:ok, error } = await db.rpc("meu_acesso_rv");
    if(error || !ok){ aviso('<b>Acesso restrito</b>O simulador não está liberado para o seu usuário. Peça ao responsável pelo painel.'); return; }
    MEU_COD = perfil && perfil.codigo_t ? String(perfil.codigo_t).toUpperCase() : null;
    if(!MEU_COD){ aviso('<b>Código T não cadastrado</b>Seu login ainda não tem código T vinculado. Peça ao responsável pelo painel para informar o seu código na tela de usuários.'); return; }
  }
  // salário lembrado: uma chave por login; a chave antiga (única para todos) é migrada para o dono ou apagada
  const email = String(session.user && session.user.email || "").toLowerCase();
  CHAVE_FIXO = "rv-fixo:"+email;
  try{ const antigo=localStorage.getItem("rv-fixo"); if(antigo!=null){ if(SOU_DONO && !localStorage.getItem(CHAVE_FIXO)) localStorage.setItem(CHAVE_FIXO, antigo); localStorage.removeItem("rv-fixo"); } }catch(e){}
  rvFixo = lerFixo();
  try{
    const [a,g,m]=await Promise.all([lerTudo("oportunidades_atuais"), lerTudo("oportunidades_ganhas"), lerTudo("metas")]);
    rows=a.map(toView).filter(o=>o.p!==100);
    ganhas=g.map(r=>{ const o=toView(r); if(o.mi==null && r.carga_em){ const d=new Date(r.carga_em); o.mi=d.getFullYear()*12+d.getMonth(); } return o; });
    metas=m;
  }catch(e){ aviso('<b>Não foi possível carregar os dados</b>'+esc(e.message||String(e))); return; }
  $("rvPage").classList.remove("hidden"); $("rvCarregando").classList.add("hidden");
  rvInit(); rvRender();
})();
