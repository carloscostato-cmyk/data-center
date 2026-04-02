function esc(e){return String(e??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}
// AVISO DE SEGURANÇA: Este sistema usa autenticação client-side com hashes SHA-256.
// Para produção, implemente autenticação server-side com HTTPS.
// Hashes gerados via SHA-256. Para alterar senhas, recalcule os hashes.
async function hashSenha(senha){
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(senha));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,"0")).join("");
}
// Senhas (SHA-256): "12345" = 5994471abb01112afcc18159f6cc74b4f511b99806da59b3caf5a9c173cacfc5
// "Epp9e72*" = 3e1c671b22dfd87fa4e5c6d40ad15db475a7a4e5c72c3b3e0b9f5b7e1c2d8a9f (exemplo — recalcule o hash real)
// Para obter o hash real de uma senha, execute no console do navegador:
// hashSenha("sua_senha_aqui").then(h => console.log(h))
const users=[
  {u:"Administrador",    h:"5994471abb01112afcc18159f6cc74b4f511b99806da59b3caf5a9c173cacfc5", modulos:null},
  {u:"alexandre.lima",   h:"HASH_DA_SENHA_AQUI_RECALCULE_NO_CONSOLE",                          modulos:null},
  {u:"lucas.vitalino",   h:"5994471abb01112afcc18159f6cc74b4f511b99806da59b3caf5a9c173cacfc5", modulos:["base","fitas","desembalagem"]},
  {u:"roberto.carvalho", h:"5994471abb01112afcc18159f6cc74b4f511b99806da59b3caf5a9c173cacfc5", modulos:["base","fitas","desembalagem"]},
  {u:"david.pereira",    h:"5994471abb01112afcc18159f6cc74b4f511b99806da59b3caf5a9c173cacfc5", modulos:["inventario","base","fitas","desembalagem","corpDashboard"]},
  {u:"claudinei.carneiro",h:"5994471abb01112afcc18159f6cc74b4f511b99806da59b3caf5a9c173cacfc5",modulos:["inventario","base","fitas","desembalagem","corpDashboard"]},
  {u:"joao.goncaes",     h:"5994471abb01112afcc18159f6cc74b4f511b99806da59b3caf5a9c173cacfc5", modulos:["operacoes","desembalagem","base"]}
],campos=["local","ambiente","rack","posicao","andar","serial","fornecedor","modelo","geracao","tipo","baia","frame","cmdb","ip","hosts","so","cpu","memoria","storage","firmware_aplicada","firmware_disponivel","bios_aplicada","bios_disponivel","eol","eos","inicioGarantia","fimGarantia","pedido","cgc","dominio","cofre"];let editIndex=null,editFitaIndex=null,opEditIndex=null,dscEditIndex=null;const form=document.getElementById("form"),tabela=document.getElementById("tabela"),auditoria=document.getElementById("auditoria"),baseConhecimento=document.getElementById("baseConhecimento"),loginDiv=document.getElementById("login"),menu=document.getElementById("menu"),painel=document.getElementById("painel"),userInfo=document.getElementById("userInfo"),logTabela=document.getElementById("logTabela"),btnLimparLogs=document.getElementById("btnLimparLogs"),fitasBackup=document.getElementById("fitasBackup"),roboFitas=document.getElementById("roboFitas"),desembalagem=document.getElementById("desembalagem"),operacoesSec=document.getElementById("operacoes"),connmapSec=document.getElementById("connmapSec"),getData=()=>JSON.parse(localStorage.getItem("inventario")||"[]"),saveData=e=>localStorage.setItem("inventario",JSON.stringify(e)),getLogs=()=>JSON.parse(localStorage.getItem("logs")||"[]"),saveLogs=e=>localStorage.setItem("logs",JSON.stringify(e)),getFitas=()=>JSON.parse(localStorage.getItem("fitas")||"[]"),saveFitas=e=>localStorage.setItem("fitas",JSON.stringify(e)),zerarHora=e=>((e=new Date(e)).setHours(0,0,0,0),e),calcularStatus=e=>e?zerarHora(e)<zerarHora(new Date)?"❌ Fora Garantia":"✅ Em Garantia":"⚠️ Sem Garantia";function registrarLog(e,t=""){const o=getLogs();o.unshift({usuario:localStorage.getItem("usuarioLogado")||"—",acao:e,detalhes:t,data:(new Date).toLocaleString()}),saveLogs(o),renderLogs()}function renderLogs(){const e=getLogs();logTabela.innerHTML="";const t=["Administrador","admin"].includes(localStorage.getItem("usuarioLogado")||"");e.forEach(e=>{const t=document.createElement("tr");t.innerHTML=`<td>${esc(e.usuario)}</td><td>${esc(e.acao)}</td><td>${esc(e.detalhes)}</td><td>${esc(e.data)}</td>`,logTabela.appendChild(t)}),btnLimparLogs.style.display=e.length&&t?"inline-block":"none";const o=document.getElementById("btnExportLogs");o&&(o.style.display=e.length?"inline-block":"none")}function exportarLogsExcel(){const e=getLogs();if(!e.length)return void alert("Nenhum log para exportar.");const t=XLSX.utils.json_to_sheet(e.map(e=>({"Usuário":e.usuario,"Ação":e.acao,Detalhes:e.detalhes,Data:e.data}))),o=XLSX.utils.book_new();XLSX.utils.book_append_sheet(o,t,"Auditoria"),XLSX.writeFile(o,"auditoria_logs.xlsx")}function limparLogs(){["Administrador","admin"].includes(localStorage.getItem("usuarioLogado")||"")?(localStorage.removeItem("logs"),renderLogs(),registrarLog("Limpeza de Logs")):alert("Apenas administradores podem limpar os logs.")}function applyTheme(e){document.body.classList.toggle("dark","dark"===e),localStorage.setItem("theme",e)}function toggleDark(){applyTheme(document.body.classList.toggle("dark")?"dark":"light")}async function fazerLogin(){const e=document.getElementById("user").value.trim(),t=document.getElementById("pass").value;if(!e||!t)return void alert("Preencha usuário e senha.");const h=await hashSenha(t),o=users.find(o=>o.u===e&&o.h===h);o?(localStorage.setItem("usuarioLogado",o.u),registrarLog("Login"),iniciarSessao(o.u)):alert("Usuário ou senha inválidos!");}function logout(){localStorage.removeItem("usuarioLogado"),location.reload()}function iniciarSessao(e){
  loginDiv&&(loginDiv.style.display="none");
  menu.style.display="block";
  // Carrega PDFs salvos no localStorage ao iniciar sessão
  setTimeout(_pdfRenderAll, 200);
  userInfo.textContent=e;
  document.body.classList.remove("login-bg");
  aplicarPermissoes(e);
}
function aplicarPermissoes(u){
  const usr=users.find(x=>x.u===u);
  const modulos=usr&&usr.modulos?usr.modulos:null;
  const map={
    menuInventario:"inventario",
    menuAuditoria:"auditoria",
    menuBase:"base",
    menuFitas:"fitas",
    menuDesembalagem:"desembalagem",
    menuOperacoes:"operacoes",
    menuDescarte:"descarte",
    menuDashboard:"corpDashboard"
  };
  Object.entries(map).forEach(([cardId,mod])=>{
    const el=document.getElementById(cardId);
    if(!el)return;
    el.style.display=(!modulos||modulos.includes(mod))?"":"none";
  });
}function abrirModulo(e){menu.style.display="none",painel.style.display="block",mostrarAba(e)}function abrirPainel(){abrirModulo("inventario")}function abrirAuditoria(){abrirModulo("auditoria")}function abrirBase(){abrirModulo("base")}function abrirFitas(){abrirModulo("fitas")}function abrirRobo(){abrirModulo("robo")}window.abrirRoboMenu=function(){document.getElementById("menu").style.display="none";document.getElementById("painel").style.display="block";mostrarAba("robo");const sub=document.getElementById("painelSubtitle");if(sub)sub.textContent="Robô de Fitas";};function abrirConnMap(){abrirModulo("connmap")}function abrirDesembalagem(){abrirModulo("desembalagem")}function abrirOperacoes(){abrirModulo("operacoes")}function abrirDescarte(){abrirModulo("descarte")}function voltarMenu(){painel.style.display="none",menu.style.display="block"}function mostrarAba(e){const t=document.getElementById("operacoes_search"),o=document.getElementById("inventario"),a=document.getElementById("descarte");o&&(o.style.display="none"),a&&(a.style.display="none"),auditoria.style.display="none",baseConhecimento.style.display="none",fitasBackup.style.display="none",roboFitas&&(roboFitas.style.display="none"),connmapSec&&(connmapSec.style.display="none"),desembalagem.style.display="none",operacoesSec&&(operacoesSec.style.display="none"),t&&(t.style.display="none"),document.getElementById("corpDashboard")&&(document.getElementById("corpDashboard").style.display="none"),"inventario"===e?(o&&(o.style.display="block")):"auditoria"===e?auditoria.style.display="block":"base"===e?baseConhecimento.style.display="block":"fitas"===e?(fitasBackup.style.display="block",renderFitas()):"robo"===e?(roboFitas&&(roboFitas.style.display="block"),roboInit()):"connmap"===e?(connmapSec&&(connmapSec.style.display="block"),cmInit()):"desembalagem"===e?(desembalagem.style.display="block",desembRenderEntrada(),desembRenderSaida(),desembRenderColeta()):"operacoes"===e&&operacoesSec?(operacoesSec.style.display="block",t&&(t.style.display="block"),opRender(),opRenderPDFs()):"descarte"===e&&(a&&(a.style.display="block",descarteRender()),document.getElementById("painelSubtitle").textContent="Descarte"),"corpDashboard"===e&&(document.getElementById("corpDashboard").style.display="block",renderCorpDashboard())}function add(){if(!document.getElementById("serial").value.trim())return void alert("O campo Serial é obrigatório!");const e=getData(),t={};campos.forEach(e=>t[e]=document.getElementById(e).value),t.status=calcularStatus(t.fimGarantia),null!==editIndex?(e[editIndex]=t,registrarLog("Edição","Registro atualizado"),editIndex=null):(e.push(t),registrarLog("Inclusão","Novo servidor")),saveData(e),atualizarTabela(),atualizarDashboard(),form.reset()}function filtrarTabela(e,t){t=(t||"").toLowerCase();const o="TABLE"===e.tagName?e.tBodies[0].rows:e.rows;Array.from(o).forEach(e=>{e.style.display=Array.from(e.cells).some(e=>(e.textContent||"").toLowerCase().includes(t))?"":"none"})}function filtrar(e){filtrarTabela(tabela,e)}function atualizarTabela(){const e=getData(),t=tabela.tBodies[0];t.innerHTML="",e.forEach((e,o)=>{const a=document.createElement("tr");campos.forEach((t,idx)=>{const o=document.createElement("td");o.textContent=e[t]||"",a.appendChild(o);
  // Inject dias p/ vencimento right after fimGarantia (index 26)
  if(t==="fimGarantia"){
    const diasTd=document.createElement("td");
    if(e.fimGarantia&&e.fimGarantia.trim()){
      const fim=new Date(e.fimGarantia.trim());
      if(!isNaN(fim)){
        const hoje=new Date();hoje.setHours(0,0,0,0);fim.setHours(0,0,0,0);
        const diff=Math.round((fim-hoje)/(1000*60*60*24));
        if(diff>30){diasTd.textContent=diff+" dias";diasTd.style.cssText="color:#16a34a;font-weight:600;white-space:nowrap";}
        else if(diff>0){diasTd.textContent=diff+" dias ⚠️";diasTd.style.cssText="color:#f59e0b;font-weight:600;white-space:nowrap";}
        else if(diff===0){diasTd.textContent="Vence hoje!";diasTd.style.cssText="color:#dc2626;font-weight:700;white-space:nowrap";}
        else{diasTd.textContent=Math.abs(diff)+" dias atrás";diasTd.style.cssText="color:#dc2626;font-weight:600;white-space:nowrap";}
      }else{diasTd.textContent="—";}
    }else{diasTd.textContent="—";}
    a.appendChild(diasTd);
  }
});
const n=document.createElement("td");n.textContent=e.status||"",a.appendChild(n);const r=document.createElement("td");r.innerHTML=`<div class="tbl-actions"><button class="btn-edit" onclick="editar(${o})" type="button">✏️ Editar</button><button class="btn-del" onclick="excluir(${o})" type="button">🗑️ Excluir</button></div>`,a.appendChild(r),t.appendChild(a)})}function editar(e){const t=getData();editIndex=e,campos.forEach(o=>document.getElementById(o).value=t[e][o]||"");const o=document.getElementById("formularioWrapper"),a=document.getElementById("btnToggleForm");o.style.display="block",a.textContent="✖ Fechar Formulário",o.scrollIntoView({behavior:"smooth",block:"start"})}function excluir(e){if(!confirm("Deseja excluir este registro?"))return;const t=getData();t.splice(e,1),saveData(t),registrarLog("Exclusão","Registro removido"),atualizarTabela(),atualizarDashboard()}function atualizarDashboard(){const e=getData(),t=e.filter(e=>"✅ Em Garantia"===e.status).length,o=e.filter(e=>"❌ Fora Garantia"===e.status).length,a=e.filter(e=>e.eol&&""!==e.eol.trim()).length,n=e.filter(e=>e.eos&&""!==e.eos.trim()).length;document.getElementById("dashTotal")&&(document.getElementById("dashTotal").textContent=e.length);document.getElementById("dashAtivos")&&(document.getElementById("dashAtivos").textContent=t);document.getElementById("dashEol")&&(document.getElementById("dashEol").textContent=a);document.getElementById("dashEos")&&(document.getElementById("dashEos").textContent=n);document.getElementById("dashGarantia")&&(document.getElementById("dashGarantia").textContent=o);const r={};e.forEach(e=>{e.fornecedor&&(r[e.fornecedor]=(r[e.fornecedor]||0)+1)});if(document.getElementById("dashFornecedor"))document.getElementById("dashFornecedor").innerHTML=Object.entries(r).map(([e,t])=>`<div><strong>${esc(e)}</strong>: ${t}</div>`).join("")||"—";const cd=document.getElementById("corpDashboard");if(cd&&cd.style.display!=="none")renderCorpDashboard();}function exportar(){if("undefined"==typeof XLSX)return void alert("Biblioteca XLSX não carregada.");const e=getData(),t=XLSX.utils.json_to_sheet(e),o=XLSX.utils.book_new();XLSX.utils.book_append_sheet(o,t,"Inventario"),XLSX.writeFile(o,"inventario.xlsx"),registrarLog("Exportação","Excel gerado")}// ── PDF persistence helpers ──
function _pdfGetStored(){try{return JSON.parse(localStorage.getItem("pdf_list")||"[]")}catch(e){return[]}}
function _pdfSaveStored(list){try{localStorage.setItem("pdf_list",JSON.stringify(list))}catch(e){alert("Aviso: não foi possível salvar o PDF no armazenamento local (arquivo muito grande ou quota excedida).");}}
function _pdfRenderRow(item,idx){
  const o=document.createElement("tr");
  o.dataset.pdfIdx=idx;
  o.innerHTML=`<td>${esc(item.nome)}</td><td>${esc(item.data)}</td>
  <td style="white-space:nowrap;display:flex;gap:6px;align-items:center">
    <button class="btn-add" onclick="visualizarPDF(_pdfGetStored()[${idx}].dataUrl)" type="button">Abrir</button>
    <a href="${item.dataUrl}" download="${esc(item.nome)}" class="btn-excel" style="display:inline-flex;align-items:center;text-decoration:none;color:#fff;font-weight:bold;padding:0 12px;height:42px;border-radius:6px">⬇ Download</a>
    <button class="btn-del" onclick="excluirPDF(this,'${esc(item.nome)}',${idx})" type="button">🗑️</button>
  </td>`;
  return o;
}
function _pdfRenderAll(){
  const lista=document.getElementById("listaPDF");
  if(!lista)return;
  lista.innerHTML="";
  _pdfGetStored().forEach((item,idx)=>lista.appendChild(_pdfRenderRow(item,idx)));
}
function salvarPDF(){
  const file=document.getElementById("pdfUpload").files[0];
  if(!file)return void alert("Selecione um PDF!");
  const reader=new FileReader();
  reader.onload=function(ev){
    const dataUrl=ev.target.result;
    const nome=file.name;
    const data=(new Date).toLocaleDateString();
    const list=_pdfGetStored();
    list.push({nome,dataUrl,data});
    _pdfSaveStored(list);
    _pdfRenderAll();
    document.getElementById("pdfUpload").value="";
    registrarLog("PDF","Documento adicionado: "+nome);
  };
  reader.readAsDataURL(file);
}function visualizarPDF(e){
  const modal=document.getElementById("modalPDF");
  const viewer=document.getElementById("pdfViewer");
  // Tenta iframe; se bloqueado, abre em nova aba
  try{
    viewer.src=e;
    viewer.dataset.url=e;
    modal.style.display="block";
    // Fallback: se iframe ficar vazio após 1.5s, abre em nova aba
    viewer.onload=function(){
      try{if(!viewer.contentDocument&&!viewer.contentWindow){window.open(e,'_blank');}}catch(err){window.open(e,'_blank');}
    };
  }catch(err){window.open(e,'_blank');}
}function abrirPDFNovaAba(){const u=document.getElementById("pdfViewer").dataset.url;if(u)window.open(u,"_blank");}
function fecharPDF(){document.getElementById("modalPDF").style.display="none";const e=document.getElementById("pdfViewer"),t=e.dataset.url;e.src="",t?.startsWith("blob:")&&URL.revokeObjectURL(t),e.dataset.url=""}function excluirPDF(e,t,idx){if(!confirm("Deseja excluir este PDF?"))return;const list=_pdfGetStored();if(!isNaN(idx)&&idx>=0&&idx<list.length){list.splice(idx,1);_pdfSaveStored(list);}registrarLog("PDF Excluído","Documento removido: "+t);_pdfRenderAll();}function addFita(){if(!document.getElementById("fitaCodigo").value.trim())return void alert("O campo Código da Fita é obrigatório!");const e=getFitas(),t={gaveta:document.getElementById("fitagaveta").value,codigo:document.getElementById("fitaCodigo").value,tipo:document.getElementById("fitaTipo").value,capacidade:document.getElementById("fitaCapacidade").value,local:document.getElementById("fitaLocal").value,status:document.getElementById("fitaStatus").value,dataRegistro:(new Date).toISOString()};null!==editFitaIndex?(t.dataRegistro=e[editFitaIndex].dataRegistro,e[editFitaIndex]=t,registrarLog("Fitas Backup","Fita editada: "+t.codigo),editFitaIndex=null):(e.push(t),registrarLog("Fitas Backup","Fita adicionada: "+t.codigo)),saveFitas(e),document.getElementById("formFitas").reset(),renderFitas()}function editarFita(e){const t=getFitas()[e];editFitaIndex=e,document.getElementById("fitagaveta").value=t.gaveta||"",document.getElementById("fitaCodigo").value=t.codigo||"",document.getElementById("fitaTipo").value=t.tipo||"",document.getElementById("fitaCapacidade").value=t.capacidade||"",document.getElementById("fitaLocal").value=t.local||"",document.getElementById("fitaStatus").value=t.status||"",registrarLog("Fitas Backup","Edição iniciada: "+(t.codigo||""));const fw=document.getElementById("formFitasWrapper"),btn=document.getElementById("btnToggleFitas");if(fw&&fw.style.display==="none"||fw&&fw.style.display===""){fw.style.display="block";if(btn)btn.textContent="✖ Fechar";}fw&&fw.scrollIntoView({behavior:"smooth",block:"start"});}function excluirFita(e){if(!confirm("Deseja excluir esta fita?"))return;const t=getFitas();t.splice(e,1),saveFitas(t),registrarLog("Fitas Backup","Fita removida"),renderFitas()}function renderFitas(){const e=document.getElementById("tabelaFitas");e.innerHTML="",getFitas().forEach((t,o)=>{const a=document.createElement("tr");a.innerHTML=`<td>${esc(t.gaveta)}</td><td>${esc(t.codigo)}</td><td>${esc(t.tipo)}</td>\n      <td>${esc(t.capacidade)}</td><td>${esc(t.local)}</td><td>${esc(t.status)}</td>\n      <td><div class="tbl-actions"><button class="btn-edit" onclick="editarFita(${o})" type="button">✏️ Editar</button><button class="btn-del" onclick="excluirFita(${o})" type="button">🗑️ Excluir</button></div></td>`,e.appendChild(a)})}function exportarFitasExcel(){const e=getFitas();if(!e.length)return void alert("Nenhuma fita para exportar.");const t=XLSX.utils.json_to_sheet(e.map(e=>({Gaveta:e.gaveta||"","Código":e.codigo||"",Tipo:e.tipo||"",Capacidade:e.capacidade||"","Localização":e.local||"",Status:e.status||"","Data Registro":e.dataRegistro||""}))),o=XLSX.utils.book_new();XLSX.utils.book_append_sheet(o,t,"Fitas de Backup"),XLSX.writeFile(o,"fitas_backup.xlsx")}function filtrarFitas(e){filtrarTabela(document.getElementById("tabelaFitas"),e)}
function gerarDeclaracaoTransporte(){
  // Pre-populate nome from logged user
  const usuario=localStorage.getItem("usuarioLogado")||"";
  const nomeEl=document.getElementById("decl_nome_completo");
  if(nomeEl&&!nomeEl.value)nomeEl.value=usuario;
  // Set today's date
  const dataEl=document.getElementById("decl_data_saida");
  if(dataEl&&!dataEl.value)dataEl.value=new Date().toISOString().split("T")[0];
  // Init lines if empty
  if(!window._declLinhas||window._declLinhas.length===0){
    window._declLinhas=[];
    declAdicionarLinha();declAdicionarLinha();
  }
  declRenderTabela();
  declAtualizarStatus();
  document.getElementById("modalDeclaracao").classList.add("open");
  registrarLog("Fitas Backup","Declaração de Transportes aberta");
}
function fecharModalDeclaracao(){
  document.getElementById("modalDeclaracao").classList.remove("open");
}
function declSetOperacao(op){
  window._declOperacao=op;
  document.getElementById("decl-btn-entrada").classList.toggle("active",op==="entrada");
  document.getElementById("decl-btn-saida").classList.toggle("active",op==="saida");
}
function declAdicionarLinha(){
  if(!window._declLinhas)window._declLinhas=[];
  window._declLinhas.push({id:Date.now()+Math.random(),barcode:"",lacre:""});
  declRenderTabela();
}
function declRemoverLinha(id){
  window._declLinhas=(window._declLinhas||[]).filter(l=>l.id!==id);
  declRenderTabela();
  declAtualizarStatus();
}
function declRenderTabela(){
  const tbody=document.getElementById("decl-material-tbody");
  if(!tbody)return;
  tbody.innerHTML="";
  (window._declLinhas||[]).forEach((l,i)=>{
    const tr=document.createElement("tr");
    tr.innerHTML=`<td class="decl-tr-num">${i+1}</td>
      <td><input type="text" value="${l.barcode}" placeholder="Ex: 000001L8" oninput="window._declLinhas[${i}].barcode=this.value;declAtualizarStatus()"></td>
      <td><input type="text" value="${l.lacre}" placeholder="Nº do lacre" oninput="window._declLinhas[${i}].lacre=this.value"></td>
      <td style="text-align:center;vertical-align:middle"><button class="decl-btn-del-row" onclick="declRemoverLinha(${l.id})" type="button" title="Remover">✕</button></td>`;
    tbody.appendChild(tr);
  });
}
function declAtualizarStatus(){
  const filled=(window._declLinhas||[]).filter(l=>l.barcode&&l.barcode.trim()).length;
  const el=document.getElementById("decl-status-count");
  if(el)el.textContent=filled+" item(s) com barcode";
}
function declGerarPDF(){
  document.getElementById("declOverlay").classList.add("show");
}
function declConfirmarPDF(){
  document.getElementById("declOverlay").classList.remove("show");
  declBuildPDF();
}
function declV(id){return(document.getElementById(id)?.value||"").trim();}
function declGetLocalidade(){return document.getElementById("decl_localidade")?.value||"";}
function declBuildPDF(){
  if(typeof window.jspdf==="undefined"){alert("Biblioteca jsPDF não carregada. Verifique sua conexão.");return;}
  const{jsPDF}=window.jspdf;
  const doc=new jsPDF({orientation:"portrait",unit:"mm",format:"a4"});
  const W=210,H=297,ml=12,mr=12,mt=10,cw=W-ml-mr;
  let y=mt;
  const operacao=window._declOperacao||"saida";
  function ln(n=5){y+=n;}
  function line(x1,y1,x2,y2,color="#cccccc",lw=0.3){doc.setDrawColor(color);doc.setLineWidth(lw);doc.line(x1,y1,x2,y2);}
  function rect(x,y2,w,h,fill,stroke){if(fill){doc.setFillColor(fill);doc.rect(x,y2,w,h,"F");}if(stroke){doc.setDrawColor(stroke);doc.setLineWidth(0.3);doc.rect(x,y2,w,h,"S");}}
  function text(txt,x,y2,opts={}){doc.setFont("helvetica",opts.style||"normal");doc.setFontSize(opts.size||9);doc.setTextColor(opts.color||"#111111");doc.text(String(txt||""),x,y2,{align:opts.align||"left",maxWidth:opts.maxWidth});}
  function cell(label,val,x,y2,w,h,opts={}){rect(x,y2,w,h,opts.bg||null,"#aaaaaa");doc.setFont("helvetica","bold");doc.setFontSize(7);doc.setTextColor("#444444");doc.text(label.toUpperCase(),x+2,y2+3.5);doc.setFont("helvetica",opts.valBold?"bold":"normal");doc.setFontSize(opts.valSize||9);doc.setTextColor(opts.valColor||"#111111");doc.text(String(val||""),x+2,y2+8,{maxWidth:w-4});}
  // HEADER
  // _claroLogoB64 replaced
  const _claroLogoB64_unused="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";
  // Logo Claro Empresas no cabeçalho do PDF
  rect(ml,y,cw,22,"#ffffff","#cccccc");
  try{doc.addImage("data:image/png;base64,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","PNG",ml+1,y+6,58,9.7);}catch(e){}
  rect(ml+60,y,cw-60,10,"#880000",null);
  doc.setFont("helvetica","bold");doc.setFontSize(13);doc.setTextColor("#ffffff");
  doc.text("ENTRADA | SAÍDA DE FITAS LTO",ml+62,y+7.5);
  rect(ml+60,y+10,65,12,"#f5f5f5","#aaaaaa");
  text("LOCALIDADE:",ml+62,y+15.5,{style:"bold",size:7,color:"#444"});
  text(declGetLocalidade(),ml+62,y+19.5,{style:"bold",size:8});
  rect(ml+125,y+10,cw-125,12,"#f5f5f5","#aaaaaa");
  text("ANDAR/SALA:",ml+127,y+15.5,{style:"bold",size:7,color:"#444"});
  text("4º ANDAR",ml+127,y+19.5,{style:"bold",size:8});
  rect(ml,mt+22,60,12,"#f0f0f0","#aaaaaa");
  text("ÓRGÃO EMITENTE",ml+4,mt+27.5,{style:"bold",size:7,color:"#444"});
  text("LOGÍSTICA",ml+4,mt+31.5,{style:"bold",size:9});
  rect(ml+60,mt+22,cw-60,12,"#e8e8e8","#aaaaaa");
  text("DATA CENTER CLARO EMPRESAS",ml+62,mt+30,{style:"bold",size:10,color:"#880000"});
  y=mt+34;ln(3);
  // LOCALIDADES CHECKBOXES
  const localidades=["Henri Dunant - SP","DC Lapa - SP","DC Ingleses - SP","DC Brasília - DF","DC Vitória - ES","DC Mackenzie - RJ"];
  const selectedLoc=declGetLocalidade();let lx=ml;
  doc.setFontSize(8);
  localidades.forEach(loc=>{
    const checked=loc===selectedLoc;
    rect(lx,y,3.5,3.5,checked?"#cc0000":"#ffffff","#666666");
    if(checked){doc.setTextColor("#ffffff");doc.setFontSize(7);doc.text("✓",lx+0.5,y+3);}
    doc.setTextColor("#111");doc.setFontSize(8);doc.text(loc,lx+5,y+3.2);lx+=32;
  });ln(7);
  const tipos=["Entrada de materiais","Saída de materiais"];lx=ml;
  tipos.forEach(t=>{
    const checked=(t==="Entrada de materiais"&&operacao==="entrada")||(t==="Saída de materiais"&&operacao==="saida");
    rect(lx,y,3.5,3.5,checked?"#cc0000":"#ffffff","#666666");
    if(checked){doc.setTextColor("#ffffff");doc.setFontSize(7);doc.text("✓",lx+0.5,y+3);}
    doc.setTextColor("#111");doc.setFontSize(8.5);doc.text(t,lx+5,y+3.2);lx+=50;
  });ln(8);
  // REMETENTE
  rect(ml,y,cw,5,"#1a1a2e",null);text("REMETENTE / EMITENTE",ml+2,y+3.8,{style:"bold",size:8,color:"#ffffff"});ln(5);
  const rh=10;
  cell("Nome / Razão Social",declV("decl_razao_social"),ml,y,cw-48,rh,{valBold:true});
  cell("C.N.P.J/MF",declV("decl_cnpj"),ml+cw-48,y,48,rh,{valBold:true});ln(rh);
  cell("Endereço",declV("decl_endereco"),ml,y,cw-80,rh);
  cell("Bairro",declV("decl_bairro"),ml+cw-80,y,46,rh);
  cell("CEP",declV("decl_cep"),ml+cw-34,y,34,rh);ln(rh);
  cell("Município",declV("decl_municipio"),ml,y,50,rh);
  cell("Fone / Fax",declV("decl_fone"),ml+50,y,45,rh);
  cell("UF",declV("decl_uf"),ml+95,y,20,rh);
  cell("Contato",declV("decl_contato"),ml+115,y,cw-115,rh);ln(rh+3);
  // SOLICITANTE
  rect(ml,y,cw,5,"#1a1a2e",null);text("DADOS DO SOLICITANTE | PERMISSIONÁRIO",ml+2,y+3.8,{style:"bold",size:8,color:"#ffffff"});ln(5);
  cell("Nome Completo",declV("decl_nome_completo"),ml,y,cw-55,rh,{valBold:true});
  cell("Identidade",declV("decl_identidade"),ml+cw-55,y,55,rh);ln(rh);
  cell("E-mail",declV("decl_email"),ml,y,cw/2,rh);
  cell("Telefone",declV("decl_telefone"),ml+cw/2,y,cw/2,rh);ln(rh);
  cell("Número do Chamado",declV("decl_chamado"),ml,y,cw/2,rh);
  let ds=declV("decl_data_saida");
  if(ds){const p=ds.split("-");ds=`${p[2]}/${p[1]}/${p[0]}`;}
  cell("Data de Saída",ds,ml+cw/2,y,cw/2,rh);ln(rh);
  cell("Observações",declV("decl_observacoes"),ml,y,cw,12);ln(15);
  // MATERIAL
  rect(ml,y,cw,5,"#1a1a2e",null);text("DADOS DO MATERIAL",ml+2,y+3.8,{style:"bold",size:8,color:"#ffffff"});ln(5);
  rect(ml,y,cw/2,6,"#1a1a2e",null);text("BARCODE",ml+cw/4,y+4.2,{style:"bold",size:8,color:"#ffffff",align:"center"});
  rect(ml+cw/2,y,cw/2,6,"#2a2a3e",null);text("LACRE",ml+cw/2+cw/4,y+4.2,{style:"bold",size:8,color:"#ffffff",align:"center"});ln(6);
  const materiais=(window._declLinhas||[]).filter(l=>l.barcode||l.lacre);
  if(materiais.length===0){
    for(let i=0;i<2;i++){rect(ml,y,cw/2,7,i%2===0?"#ffffff":"#f9f9f9","#cccccc");rect(ml+cw/2,y,cw/2,7,i%2===0?"#ffffff":"#f9f9f9","#cccccc");ln(7);}
  }else{
    materiais.forEach((l,i)=>{
      const bg=i%2===0?"#ffffff":"#f5f5f5";
      rect(ml,y,cw/2,7,bg,"#cccccc");rect(ml+cw/2,y,cw/2,7,bg,"#cccccc");
      text(l.barcode,ml+cw/4,y+4.8,{size:9,align:"center"});
      text(l.lacre,ml+cw/2+cw/4,y+4.8,{size:9,align:"center"});ln(7);
    });
  }ln(4);
  // DECLARAÇÃO
  const nome=declV("decl_nome_completo")||"___________________";
  const rg=declV("decl_identidade")||"___________________";
  const opStr=operacao==="saida"?"sair":"entrar em";
  const decl2=`Declaramos que, ${nome} portador do documento RG nº ${rg} está autorizado(a) a ${opStr} deste edifício, portando os equipamentos descritos neste documento. O portador está ciente de que deverá entregar esta autorização na portaria de saída, e que estará sujeito a vistoria.`;
  doc.setFont("helvetica","normal");doc.setFontSize(9);doc.setTextColor("#111");
  const lines2=doc.splitTextToSize(decl2,cw);
  doc.text(lines2,ml,y,{maxWidth:cw});y+=lines2.length*4.5;ln(4);
  let ds2=declV("decl_data_saida"),dia="___",mes="_______________",ano=new Date().getFullYear().toString();
  if(ds2){const p=ds2.split("-");dia=p[2];ano=p[0];const meses=["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];mes=meses[parseInt(p[1])-1];}
  text(`São Paulo, ${dia} de ${mes} de ${ano}`,ml,y,{size:9});ln(12);
  line(ml+80,y,ml+cw,y,"#888888",0.4);
  text("Assinatura do Portador",ml+(cw/2+80/2),y+4,{size:8,color:"#666",align:"center"});ln(14);
  const bw=(cw-8)/3;
  ["SEGURANÇA","LOGÍSTICA | DATA CENTER","RECEPÇÃO"].forEach((b,i)=>{
    const bx=ml+i*(bw+4);
    line(bx,y,bx+bw,y,"#888888",0.4);
    text(b,bx+bw/2,y+4.5,{style:"bold",size:8,color:"#555",align:"center"});
  });ln(12);
  // ENDEREÇOS
  if(y>H-50){doc.addPage();y=mt;}
  rect(ml,y,cw,5,"#1a1a2e",null);text("ENDEREÇOS LOCALIDADES DATA CENTER",ml+2,y+3.8,{style:"bold",size:8,color:"#ffffff"});ln(5);
  const addrH=["SITE","ENDEREÇO","Nº","COMPLEMENTO","BAIRRO","CEP","CIDADE","UF","TEL."];
  const addrW=[32,36,10,18,22,18,20,8,22];
  addrH.forEach((h,i)=>{const ax=ml+addrW.slice(0,i).reduce((a,b)=>a+b,0);rect(ax,y,addrW[i],5.5,"#1a1a2e",null);text(h,ax+1,y+4,{style:"bold",size:6.5,color:"#ffffff"});});ln(5.5);
  const _savedAddrs = (()=>{try{return JSON.parse(localStorage.getItem('decl_addr_localidades'))||null;}catch(e){return null;}})();
  const addrs=(_savedAddrs||[
    {site:"DATA CENTER BRASÍLIA",endereco:"SCSQ 5 ED. EMBRATEL BL D",numero:"5",complemento:"",bairro:"ASA SUL",cep:"70328-900",cidade:"BRASÍLIA",uf:"DF",telefone:"(61) 2106-8246"},
    {site:"SEDE HENRI DUNAT",endereco:"R HENRI DUNANT",numero:"780",complemento:"",bairro:"SANTO AMARO",cep:"04.709-110",cidade:"SAO PAULO",uf:"SP",telefone:"(11) 4313-4620"},
    {site:"DATA CENTER LAPA",endereco:"RUA ALDO DE AZEVEDO",numero:"200",complemento:"2º ANDAR",bairro:"VILA MADALENA",cep:"05453-030",cidade:"SÃO PAULO",uf:"SP",telefone:"(11) 2121-3874"},
    {site:"DATA CENTER INGLESES",endereco:"RUA DOS INGLESES",numero:"600",complemento:"4º ANDAR",bairro:"M. DOS INGLESES",cep:"01329-904",cidade:"SÃO PAULO",uf:"SP",telefone:"(11) 2121-3874"},
    {site:"DATA CENTER VITÓRIA",endereco:"AV. JERONIMO MONTEIRO",numero:"174",complemento:"",bairro:"CENTRO",cep:"29010-002",cidade:"ESPÍRITO SANTO",uf:"ES",telefone:"(11) 2121-3874"},
    {site:"DATA CENTER MACKENZIE",endereco:"RUA SENADOR POMPEU",numero:"119",complemento:"12º ANDAR",bairro:"CENTRO",cep:"20221-291",cidade:"RIO DE JANEIRO",uf:"RJ",telefone:"(21) 2121-2414"},
  ]).map(r=>[r.site||'',r.endereco||'',r.numero||'',r.complemento||'',r.bairro||'',r.cep||'',r.cidade||'',r.uf||'',r.telefone||'']);
  addrs.forEach((row,ri)=>{
    const rh2=7;
    row.forEach((c2,ci)=>{
      const ax=ml+addrW.slice(0,ci).reduce((a,b)=>a+b,0);
      rect(ax,y,addrW[ci],rh2,ri%2===0?"#ffffff":"#f5f5f5","#cccccc");
      text(c2,ax+1.5,y+4.5,{size:7,color:ci===0?"#cc0000":"#111"});
    });ln(rh2);
  });
  ln(4);
  text("Claro Empresas — catia.santos@globalhitss.com.br | ricardo.moraes@globalhitss.com.br | heitor.aoki@claro.com.br",W/2,y,{size:7,color:"#888",align:"center"});ln(4);
  text("(11) 2121-2022 | (11) 2121-3647 | (11) 2121-2886",W/2,y,{size:7,color:"#888",align:"center"});
  const fname=`Declaracao_Fitas_${(declV("decl_nome_completo")||"sem_nome").replace(/\s+/g,"_")}_${new Date().toISOString().slice(0,10)}.pdf`;
  doc.save(fname);
  registrarLog("Fitas Backup","Declaração de Transportes PDF gerado");
}
// Init decl state
window._declOperacao='saida';
window._declLinhas=[];
// Close modal on backdrop click
document.addEventListener('click',function(e){
  const modal=document.getElementById('modalDeclaracao');
  if(e.target===modal)fecharModalDeclaracao();
});const getDesembEntrada=()=>JSON.parse(localStorage.getItem("desemb_entrada")||"[]"),saveDesembEntrada=e=>localStorage.setItem("desemb_entrada",JSON.stringify(e)),getDesembSaida=()=>JSON.parse(localStorage.getItem("desemb_saida")||"[]"),saveDesembSaida=e=>localStorage.setItem("desemb_saida",JSON.stringify(e)),getDesembColeta=()=>JSON.parse(localStorage.getItem("desemb_coleta")||"[]"),saveDesembColeta=e=>localStorage.setItem("desemb_coleta",JSON.stringify(e));function desembToggleArmario(){document.getElementById("de_armario").style.display="ARMARIO"===document.getElementById("de_setor").value?"block":"none"}function exportarDesembEntradaExcel(){const e=getDesembEntrada();if(!e.length)return void alert("Nenhum registro de entrada para exportar.");const t=XLSX.utils.json_to_sheet(e.map(e=>({Caixa:e.caixa||"",Setor:e.setor||"","Armário":e.armario||"",Cliente:e.cliente||"","Nota Fiscal":e.nota||"","RMA/W.O":e.rma||"","Chave de Acesso":e.chave||"",Volumes:e.volumes||"","Descrição":e.descricao||"","Responsável":e.responsavel||"",Data:e.data||""}))),o=XLSX.utils.book_new();XLSX.utils.book_append_sheet(o,t,"Entrada"),XLSX.writeFile(o,"desembalagem_entrada.xlsx")}function exportarDesembSaidaExcel(){const e=getDesembSaida();if(!e.length)return void alert("Nenhum registro de saída para exportar.");const t=XLSX.utils.json_to_sheet(e.map(e=>({Caixa:e.caixa||"","Responsável":e.responsavel||"",Motivo:e.motivo||"",Data:e.data||""}))),o=XLSX.utils.book_new();XLSX.utils.book_append_sheet(o,t,"Saída"),XLSX.writeFile(o,"desembalagem_saida.xlsx")}function exportarDesembColetaExcel(){const e=getDesembColeta();if(!e.length)return void alert("Nenhum registro de coleta para exportar.");const t=XLSX.utils.json_to_sheet(e.map(e=>({Caixa:e.caixa||"","Responsável":e.responsavel||"","Situação":e.situacao||"",Data:e.data||""}))),o=XLSX.utils.book_new();XLSX.utils.book_append_sheet(o,t,"Coleta"),XLSX.writeFile(o,"desembalagem_coleta.xlsx")}function desembSalvarEntrada(){const e=document.getElementById("de_caixa").value.trim();if(!e)return void alert("Informe o número da caixa!");const t=document.getElementById("de_foto"),o=getDesembEntrada();o.push({caixa:e,setor:document.getElementById("de_setor").value,armario:document.getElementById("de_armario").value,cliente:document.getElementById("de_cliente").value,nota:document.getElementById("de_nota").value,rma:document.getElementById("de_rma").value,chave:document.getElementById("de_chave").value,volumes:document.getElementById("de_volumes").value,descricao:document.getElementById("de_descricao").value,responsavel:"",foto:t.files?.[0]?URL.createObjectURL(t.files[0]):"",data:(new Date).toLocaleString()}),saveDesembEntrada(o),registrarLog("Desembalagem - Entrada","Caixa: "+e),["de_caixa","de_setor","de_armario","de_cliente","de_nota","de_rma","de_chave","de_volumes","de_descricao"].forEach(e=>{const t=document.getElementById(e);t&&("SELECT"===t.tagName?t.selectedIndex=0:t.value="")}),document.getElementById("de_armario").style.display="none",desembRenderEntrada(),alert("Entrada salva com sucesso!")}function desembRenderEntrada(){const e=document.getElementById("tabelaDesembEntrada"),t=getDesembEntrada();e.innerHTML="",t.forEach((t,o)=>{const a=document.createElement("tr");const evs=t.evidencias||[];const evHtml=evs.length?evs.slice(0,3).map((f,fi)=>f.type&&f.type.startsWith('image/')?`<span class="ev-inline-thumb" onclick="evLightboxOpen('desemb_entrada',${o},${fi})" title="${esc(f.name)}"><img src="${f.data}" alt="${esc(f.name)}"></span>`:`<span class="ev-inline-thumb" onclick="evLightboxOpen('desemb_entrada',${o},${fi})" title="${esc(f.name)}">📄 ${esc(f.name.substring(0,10))}…</span>`).join(' ')+(evs.length>3?` <span style="font-size:11px;color:var(--sub)">+${evs.length-3}</span>`:''):'<span style="font-size:11px;color:var(--sub)">—</span>';a.innerHTML=`<td>${esc(t.caixa)}</td><td>${esc(t.setor)}</td><td>${esc(t.armario)}</td>\n      <td>${esc(t.cliente)}</td><td>${esc(t.nota)}</td><td>${esc(t.rma)}</td>\n      <td>${esc(t.chave)}</td><td>${esc(t.volumes)}</td><td>${esc(t.descricao)}</td>\n      <td>${esc(t.responsavel)}</td>\n      <td>${t.foto?`<button class="btn-dark" onclick="verFoto('${esc(t.foto)}')" type="button">Ver Foto</button>`:"—"}</td>\n      <td>${evHtml}</td>\n      <td><div class="tbl-actions"><button class="btn-dark" onclick="abrirEvidencias('desemb_entrada',${o})" type="button" title="Evidências">📎</button><button class="btn-edit" onclick="desembEditarEntrada(${o})" type="button">✏️ Editar</button><button class="btn-del" onclick="desembExcluirEntrada(${o})" type="button">🗑️ Excluir</button></div></td>`,e.appendChild(a)})}
function desembEditarEntrada(idx){
  const t=getDesembEntrada()[idx];if(!t)return;
  const fw=document.getElementById("formEntradaWrapper"),btn=document.getElementById("btnToggleEntrada");
  if(fw.style.display==="none"||fw.style.display===""){fw.style.display="block";btn.textContent="✖ Fechar";}
  document.getElementById("de_caixa").value=t.caixa||"";
  document.getElementById("de_setor").value=t.setor||"";
  desembToggleArmario();
  document.getElementById("de_armario").value=t.armario||"";
  document.getElementById("de_cliente").value=t.cliente||"";
  document.getElementById("de_nota").value=t.nota||"";
  document.getElementById("de_rma").value=t.rma||"";
  document.getElementById("de_chave").value=t.chave||"";
  document.getElementById("de_volumes").value=t.volumes||"";
  document.getElementById("de_descricao").value=t.descricao||"";
  const btnSalvar=document.querySelector('[onclick="desembSalvarEntrada()"]');
  if(btnSalvar){btnSalvar.textContent="💾 Salvar Alterações";btnSalvar.setAttribute("onclick","desembSalvarEdicaoEntrada("+idx+")");}
  fw.scrollIntoView({behavior:"smooth",block:"start"});
}
function desembSalvarEdicaoEntrada(idx){
  const e=document.getElementById("de_caixa").value.trim();if(!e)return void alert("Informe o número da caixa!");
  const data=getDesembEntrada();
  data[idx]={...data[idx],caixa:e,setor:document.getElementById("de_setor").value,armario:document.getElementById("de_armario").value,cliente:document.getElementById("de_cliente").value,nota:document.getElementById("de_nota").value,rma:document.getElementById("de_rma").value,chave:document.getElementById("de_chave").value,volumes:document.getElementById("de_volumes").value,descricao:document.getElementById("de_descricao").value};
  saveDesembEntrada(data);registrarLog("Desembalagem - Entrada","Editado: "+e);
  ["de_caixa","de_setor","de_armario","de_cliente","de_nota","de_rma","de_chave","de_volumes","de_descricao"].forEach(id=>{const el=document.getElementById(id);"SELECT"===el.tagName?el.selectedIndex=0:el.value="";});
  document.getElementById("de_armario").style.display="none";
  const btnSalvar=document.querySelector('[onclick^="desembSalvarEdicaoEntrada"]');
  if(btnSalvar){btnSalvar.textContent="Salvar Entrada";btnSalvar.setAttribute("onclick","desembSalvarEntrada()");}
  desembRenderEntrada();alert("Entrada atualizada!");
}
function desembExcluirEntrada(e){if(!confirm("Excluir este registro?"))return;const t=getDesembEntrada(),[o]=t.splice(e,1);if(saveDesembEntrada(t),o?.foto?.startsWith("blob:"))try{URL.revokeObjectURL(o.foto)}catch(e){}registrarLog("Desembalagem - Exclusão Entrada"),desembRenderEntrada()}function desembSalvarSaida(){const e=document.getElementById("ds_caixa").value.trim(),t=document.getElementById("ds_responsavel").value.trim(),o=document.getElementById("ds_motivo").value;if(!e)return void alert("Informe o número da caixa!");if(!t)return void alert("Informe o responsável!");if(!o)return void alert("Selecione o motivo!");const a=document.getElementById("ds_foto"),n=getDesembSaida();n.push({caixa:e,responsavel:t,motivo:o,data:(new Date).toLocaleString(),foto:a?.files?.[0]?URL.createObjectURL(a.files[0]):""}),saveDesembSaida(n),registrarLog("Desembalagem - Saída","Caixa: "+e),["ds_caixa","ds_responsavel","ds_motivo"].forEach(e=>{const t=document.getElementById(e);"SELECT"===t.tagName?t.selectedIndex=0:t.value=""}),a&&(a.value=""),desembRenderSaida(),alert("Saída registrada!")}function desembRenderSaida(){const e=document.getElementById("tabelaDesembSaida"),t=getDesembSaida();e.innerHTML="",t.forEach((t,o)=>{const a=document.createElement("tr");const evs=t.evidencias||[];const evHtml=evs.length?evs.slice(0,3).map((f,fi)=>f.type&&f.type.startsWith('image/')?`<span class="ev-inline-thumb" onclick="evLightboxOpen('desemb_saida',${o},${fi})" title="${esc(f.name)}"><img src="${f.data}" alt="${esc(f.name)}"></span>`:`<span class="ev-inline-thumb" onclick="evLightboxOpen('desemb_saida',${o},${fi})" title="${esc(f.name)}">📄 ${esc(f.name.substring(0,10))}…</span>`).join(' ')+(evs.length>3?` <span style="font-size:11px;color:var(--sub)">+${evs.length-3}</span>`:''):'<span style="font-size:11px;color:var(--sub)">—</span>';a.innerHTML=`<td>${esc(t.caixa)}</td><td>${esc(t.responsavel)}</td><td>${esc(t.motivo)}</td>\n      <td>${esc(t.data)}</td>\n      <td>${t.foto?`<button class="btn-dark" onclick="verFoto('${esc(t.foto)}')" type="button">Ver Foto</button>`:"—"}</td>\n      <td>${evHtml}</td>\n      <td><div class="tbl-actions"><button class="btn-dark" onclick="abrirEvidencias('desemb_saida',${o})" type="button" title="Evidências">📎</button><button class="btn-edit" onclick="desembEditarSaida(${o})" type="button">✏️ Editar</button><button class="btn-del" onclick="desembExcluirSaida(${o})" type="button">🗑️ Excluir</button></div></td>`,e.appendChild(a)})}function desembEditarSaida(idx){
  const t=getDesembSaida()[idx];if(!t)return;
  const fw=document.getElementById("formSaidaWrapper"),btn=document.getElementById("btnToggleSaida");
  if(fw.style.display==="none"||fw.style.display===""){fw.style.display="block";btn.textContent="✖ Fechar";}
  document.getElementById("ds_caixa").value=t.caixa||"";
  document.getElementById("ds_responsavel").value=t.responsavel||"";
  document.getElementById("ds_motivo").value=t.motivo||"";
  const btnSalvar=document.querySelector('[onclick="desembSalvarSaida()"]');
  if(btnSalvar){btnSalvar.textContent="💾 Salvar Alterações";btnSalvar.setAttribute("onclick","desembSalvarEdicaoSaida("+idx+")");}
  fw.scrollIntoView({behavior:"smooth",block:"start"});
}
function desembSalvarEdicaoSaida(idx){
  const e=document.getElementById("ds_caixa").value.trim(),t=document.getElementById("ds_responsavel").value.trim(),o=document.getElementById("ds_motivo").value;
  if(!e)return void alert("Informe o número da caixa!");if(!t)return void alert("Informe o responsável!");if(!o)return void alert("Selecione o motivo!");
  const data=getDesembSaida();
  data[idx]={...data[idx],caixa:e,responsavel:t,motivo:o};
  saveDesembSaida(data);registrarLog("Desembalagem - Saída","Editado: "+e);
  ["ds_caixa","ds_responsavel","ds_motivo"].forEach(id=>{const el=document.getElementById(id);"SELECT"===el.tagName?el.selectedIndex=0:el.value="";});
  const btnSalvar=document.querySelector('[onclick^="desembSalvarEdicaoSaida"]');
  if(btnSalvar){btnSalvar.textContent="Registrar Saída";btnSalvar.setAttribute("onclick","desembSalvarSaida()");}
  desembRenderSaida();alert("Saída atualizada!");
}
function desembExcluirSaida(e){if(!confirm("Excluir este registro?"))return;const t=getDesembSaida(),[o]=t.splice(e,1);if(saveDesembSaida(t),o?.foto?.startsWith("blob:"))try{URL.revokeObjectURL(o.foto)}catch(e){}registrarLog("Desembalagem - Exclusão Saída"),desembRenderSaida()}function desembSalvarColeta(){const e=document.getElementById("dc_caixa").value.trim(),t=document.getElementById("dc_responsavel").value.trim(),o=document.getElementById("dc_situacao").value;if(!e)return void alert("Informe o número da caixa!");if(!t)return void alert("Informe o responsável!");if(!o)return void alert("Selecione a situação!");const a=document.getElementById("dc_foto"),n=getDesembColeta();n.push({caixa:e,responsavel:t,situacao:o,data:(new Date).toLocaleString(),foto:a?.files?.[0]?URL.createObjectURL(a.files[0]):""}),saveDesembColeta(n),registrarLog("Desembalagem - Coleta","Caixa: "+e),["dc_caixa","dc_responsavel","dc_situacao"].forEach(e=>{const t=document.getElementById(e);"SELECT"===t.tagName?t.selectedIndex=0:t.value=""}),a&&(a.value=""),desembRenderColeta(),alert("Coleta registrada!")}function desembRenderColeta(){const e=document.getElementById("tabelaDesembColeta"),t=getDesembColeta();e.innerHTML="",t.forEach((t,o)=>{const a=document.createElement("tr");const evs=t.evidencias||[];const evHtml=evs.length?evs.slice(0,3).map((f,fi)=>f.type&&f.type.startsWith('image/')?`<span class="ev-inline-thumb" onclick="evLightboxOpen('desemb_coleta',${o},${fi})" title="${esc(f.name)}"><img src="${f.data}" alt="${esc(f.name)}"></span>`:`<span class="ev-inline-thumb" onclick="evLightboxOpen('desemb_coleta',${o},${fi})" title="${esc(f.name)}">📄 ${esc(f.name.substring(0,10))}…</span>`).join(' ')+(evs.length>3?` <span style="font-size:11px;color:var(--sub)">+${evs.length-3}</span>`:''):'<span style="font-size:11px;color:var(--sub)">—</span>';a.innerHTML=`<td>${esc(t.caixa)}</td><td>${esc(t.responsavel)}</td><td>${esc(t.situacao)}</td>\n      <td>${esc(t.data)}</td>\n      <td>${t.foto?`<button class="btn-dark" onclick="verFoto('${esc(t.foto)}')" type="button">Ver Foto</button>`:"—"}</td>\n      <td>${evHtml}</td>\n      <td><div class="tbl-actions"><button class="btn-dark" onclick="abrirEvidencias('desemb_coleta',${o})" type="button" title="Evidências">📎</button><button class="btn-edit" onclick="desembEditarColeta(${o})" type="button">✏️ Editar</button><button class="btn-del" onclick="desembExcluirColeta(${o})" type="button">🗑️ Excluir</button></div></td>`,e.appendChild(a)})}function desembEditarColeta(idx){
  const t=getDesembColeta()[idx];if(!t)return;
  const fw=document.getElementById("formColetaWrapper"),btn=document.getElementById("btnToggleColeta");
  if(fw.style.display==="none"||fw.style.display===""){fw.style.display="block";btn.textContent="✖ Fechar";}
  document.getElementById("dc_caixa").value=t.caixa||"";
  document.getElementById("dc_responsavel").value=t.responsavel||"";
  document.getElementById("dc_situacao").value=t.situacao||"";
  const btnSalvar=document.querySelector('[onclick="desembSalvarColeta()"]');
  if(btnSalvar){btnSalvar.textContent="💾 Salvar Alterações";btnSalvar.setAttribute("onclick","desembSalvarEdicaoColeta("+idx+")");}
  fw.scrollIntoView({behavior:"smooth",block:"start"});
}
function desembSalvarEdicaoColeta(idx){
  const e=document.getElementById("dc_caixa").value.trim(),t=document.getElementById("dc_responsavel").value.trim(),o=document.getElementById("dc_situacao").value;
  if(!e)return void alert("Informe o número da caixa!");if(!t)return void alert("Informe o responsável!");if(!o)return void alert("Selecione a situação!");
  const data=getDesembColeta();
  data[idx]={...data[idx],caixa:e,responsavel:t,situacao:o};
  saveDesembColeta(data);registrarLog("Desembalagem - Coleta","Editado: "+e);
  ["dc_caixa","dc_responsavel","dc_situacao"].forEach(id=>{const el=document.getElementById(id);"SELECT"===el.tagName?el.selectedIndex=0:el.value="";});
  const btnSalvar=document.querySelector('[onclick^="desembSalvarEdicaoColeta"]');
  if(btnSalvar){btnSalvar.textContent="Registrar Coleta";btnSalvar.setAttribute("onclick","desembSalvarColeta()");}
  desembRenderColeta();alert("Coleta atualizada!");
}
function desembExcluirColeta(e){if(!confirm("Excluir este registro?"))return;const t=getDesembColeta(),[o]=t.splice(e,1);if(saveDesembColeta(t),o?.foto?.startsWith("blob:"))try{URL.revokeObjectURL(o.foto)}catch(e){}registrarLog("Desembalagem - Exclusão Coleta"),desembRenderColeta()}function desembFiltrarEntrada(e){filtrarTabela(document.getElementById("tabelaDesembEntrada"),e)}function desembFiltrarSaida(e){filtrarTabela(document.getElementById("tabelaDesembSaida"),e)}function desembFiltrarColeta(e){filtrarTabela(document.getElementById("tabelaDesembColeta"),e)}function verFoto(e){const t=document.createElement("div");t.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.8);display:flex;align-items:center;justify-content:center;z-index:9999;",t.innerHTML=`\n    <div style="background:#fff;padding:15px;border-radius:10px;max-width:95vw;max-height:90vh;display:flex;flex-direction:column;align-items:center">\n      <img src="${esc(e)}" alt="Foto" style="max-width:90vw;max-height:70vh;object-fit:contain;border-radius:6px" />\n      <br />\n      <div style="display:flex;gap:8px">\n        <a href="${esc(e)}" download class="btn-excel"\n           style="text-decoration:none;display:inline-block;color:#fff;">Baixar</a>\n        <button class="btn-del" onclick="this.closest('[style*=fixed]').remove()" type="button">Fechar</button>\n      </div>\n    </div>`,document.body.appendChild(t)}const getOps=()=>JSON.parse(localStorage.getItem("operacoes")||"[]"),saveOps=e=>localStorage.setItem("operacoes",JSON.stringify(e));function opAdicionarPosicao(){const e=document.getElementById("op_posicoes_wrap"),t=document.createElement("div");t.style.cssText="display:flex;gap:6px;align-items:center;";const o=document.createElement("input");o.placeholder="Posição",o.className="op_posicao_extra",o.style.flex="1";const a=document.createElement("button");a.type="button",a.className="btn-clear",a.style.cssText="height:36px;padding:0 12px;flex-shrink:0",a.textContent="–",a.onclick=()=>t.remove(),t.appendChild(o),t.appendChild(a),e.appendChild(t)}function opFormReset(){["op_cliente","op_local","op_posicao","op_equip","op_fabricante","op_modelo","op_serial","op_hostname","op_obs"].forEach(e=>{const t=document.getElementById(e);t&&(t.value="")}),opEditIndex=null}function opAdd(){const e=e=>(document.getElementById(e)?.value||"").trim(),t=e("op_equip");if(!t)return void alert("Preencha pelo menos o campo Equipamento.");const o=getOps(),a={cliente:e("op_cliente"),local:e("op_local"),posicao:e("op_posicao"),equipamento:t,fabricante:e("op_fabricante"),modelo:e("op_modelo"),serial:e("op_serial"),hostname:e("op_hostname"),obs:e("op_obs"),ts:(new Date).toISOString()};null!==opEditIndex?(o[opEditIndex]?.ts&&(a.ts=o[opEditIndex].ts),o[opEditIndex]=a,registrarLog("Operações","Relação editada"),opEditIndex=null):(o.push(a),registrarLog("Operações","Relação adicionada")),saveOps(o),opFormReset(),opRender()}function opEditar(e){const t=getOps()[e];opEditIndex=e,document.getElementById("op_cliente").value=t.cliente||"",document.getElementById("op_local").value=t.local||"",document.getElementById("op_posicao").value=t.posicao||"",document.getElementById("op_equip").value=t.equipamento||"",document.getElementById("op_fabricante").value=t.fabricante||"",document.getElementById("op_modelo").value=t.modelo||"",document.getElementById("op_serial").value=t.serial||"",document.getElementById("op_hostname").value=t.hostname||"",document.getElementById("op_obs").value=t.obs||"",document.querySelectorAll(".op_posicao_extra").forEach(e=>e.remove())}function opEditarTs(e){const t=getOps().findIndex(t=>t.ts===e);t>=0&&opEditar(t)}function opExcluirTs(e){if(!confirm("Deseja excluir esta relação?"))return;const t=getOps(),o=t.findIndex(t=>t.ts===e);o<0||(t.splice(o,1),saveOps(t),registrarLog("Operações","Relação excluída"),opRender())}function exportarOperacoesExcel(){const e=getOps();if(!e.length)return void alert("Nenhum registro para exportar.");const t=XLSX.utils.json_to_sheet(e.map(e=>({Cliente:e.cliente||"",Local:e.local||"","Posição":e.posicao||"",Equipamento:e.equipamento||"",Fabricante:e.fabricante||"",Modelo:e.modelo||"",Serial:e.serial||"",Hostname:e.hostname||"","Observação":e.obs||""}))),o=XLSX.utils.book_new();XLSX.utils.book_append_sheet(o,t,"Operações"),XLSX.writeFile(o,"operacoes.xlsx")}function opRender(e=""){const t=getOps(),o=document.querySelector("#op_tabela1 tbody");o.innerHTML="";const a=(e||"").toLowerCase();(a?t.filter(e=>[e.cliente,e.local,e.posicao,e.equipamento,e.fabricante,e.modelo,e.serial,e.hostname,e.obs].some(e=>(e||"").toLowerCase().includes(a))):t).forEach(e=>{const t=document.createElement("tr");t.innerHTML=`<td>${esc(e.cliente)}</td><td>${esc(e.local)}</td><td>${esc(e.posicao)}</td><td>${esc(e.equipamento)}</td><td>${esc(e.fabricante)}</td><td>${esc(e.modelo)}</td><td>${esc(e.serial)}</td><td>${esc(e.hostname)}</td><td>${esc(e.obs)}</td>\n      <td>\n        <div class="tbl-actions"><button class="btn-edit" onclick="opEditarTs('${e.ts}')" type="button">✏️ Editar</button><button class="btn-del" onclick="opExcluirTs('${e.ts}')" type="button">🗑️ Excluir</button></div>\n      </td>`,o.appendChild(t)})}function opFiltrar(e){opRender(e)}const getOpPDFs=()=>JSON.parse(localStorage.getItem("operacoes_pdfs")||"[]"),saveOpPDFs=e=>localStorage.setItem("operacoes_pdfs",JSON.stringify(e));function opSalvarPDF(){const e=document.getElementById("op_pdf").files[0];if(!e)return void alert("Selecione um PDF!");const t=URL.createObjectURL(e),o=document.createElement("tr");o.innerHTML=`<td>${esc(e.name)}</td><td>${(new Date).toLocaleDateString()}</td>\n    <td><button class="btn-add" onclick="visualizarPDF('${t}')" type="button">Abrir</button>\n        <button class="btn-del" onclick="opExcluirPDF(this,'${esc(e.name)}','${t}')" type="button">🗑️</button></td>`,document.getElementById("op_listaPDF").appendChild(o);const a=getOpPDFs();a.unshift({nome:e.name,data:(new Date).toLocaleString()}),saveOpPDFs(a),registrarLog("Operações - PDF","Documento adicionado: "+e.name),document.getElementById("op_pdf").value=""}function opRenderPDFs(){const e=document.getElementById("op_listaPDF");e.innerHTML="",getOpPDFs().forEach(t=>{const o=document.createElement("tr");o.innerHTML=`<td>${esc(t.nome)}</td><td>${new Date(t.data).toLocaleDateString()}</td>\n      <td><em>Arquivo precisa ser reenviado para abrir</em>\n          <button class="btn-del" onclick="opExcluirPDF(this,'${esc(t.nome)}')" type="button">🗑️</button></td>`,e.appendChild(o)})}function opExcluirPDF(e,t,o){confirm("Deseja excluir este PDF?")&&(e.closest("tr").remove(),o?.startsWith("blob:")&&URL.revokeObjectURL(o),saveOpPDFs(getOpPDFs().filter(e=>e.nome!==t)),registrarLog("Operações - PDF","Documento removido: "+t))}function toggleFitas(){const e=document.getElementById("formFitasWrapper"),t=document.getElementById("btnToggleFitas");"none"===e.style.display||""===e.style.display?(e.style.display="block",t.textContent="✖ Fechar"):(e.style.display="none",t.textContent="➕ Nova Fita",document.getElementById("formFitas").reset(),editFitaIndex=null)}function toggleOp(e,t){const o=document.getElementById(e),a=document.getElementById(t),n={formOpWrapper:"➕ Adicionar Relação",formOpPdfWrapper:"➕ Inserir PDF"};"none"===o.style.display||""===o.style.display?(o.style.display="block",a.textContent="✖ Fechar"):(o.style.display="none",a.textContent=n[e]||"➕ Adicionar")}function toggleDesemb(e,t){const o=document.getElementById(e),a=document.getElementById(t);if("none"===o.style.display||""===o.style.display)o.style.display="block",a.textContent="✖ Fechar";else{o.style.display="none";const t={formEntradaWrapper:"➕ Registrar Entrada",formSaidaWrapper:"➕ Registrar Saída",formColetaWrapper:"➕ Registrar Coleta"};a.textContent=t[e]||"➕ Registrar"}}function toggleFormulario(){const e=document.getElementById("formularioWrapper"),t=document.getElementById("btnToggleForm");"none"===e.style.display||""===e.style.display?(e.style.display="block",t.textContent="✖ Fechar Formulário"):(e.style.display="none",t.textContent="➕ Novo Servidor",form.reset(),editIndex=null)}const ROBO_LIMITE=26;let roboTbody,roboBtnAdd,roboBtnRem,roboCounter,roboIniciado=!1;function roboInit(){roboIniciado||(roboTbody=document.getElementById("robo_tbody"),roboBtnAdd=document.getElementById("robo_btnAdd"),roboBtnRem=document.getElementById("robo_btnRem"),roboCounter=document.getElementById("robo_counter"),roboIniciado=!0,roboAtualizarContador(),roboAtualizarDashboard())}function roboAtualizarDashboard(){if(!roboTbody)return;let e=0,t=0,o=0,a=0;roboTbody.querySelectorAll("tr").forEach(n=>{const r=n.querySelector(".robo-col-status");if(!r)return;const i=r.value;i&&(e++,"Pendente"===i&&t++,"Concluída"===i&&o++,"Cancelada"===i&&a++)}),document.getElementById("robo_totalReg")&&(document.getElementById("robo_totalReg").textContent=e);document.getElementById("robo_pendentes")&&(document.getElementById("robo_pendentes").textContent=t);document.getElementById("robo_concluidas")&&(document.getElementById("robo_concluidas").textContent=o);document.getElementById("robo_canceladas")&&(document.getElementById("robo_canceladas").textContent=a);const cd=document.getElementById("corpDashboard");if(cd&&cd.style.display!=="none"){const rk=document.getElementById("corpRoboKpis");if(rk)rk.innerHTML='<div class="corp-kpi-card" style="--kpi-color:#7c3aed"><div class="corp-kpi-icon">🤖</div><div class="corp-kpi-value">'+e+'</div><div class="corp-kpi-label">Total Movimentações</div></div><div class="corp-kpi-card" style="--kpi-color:#f59e0b"><div class="corp-kpi-icon">⏳</div><div class="corp-kpi-value">'+t+'</div><div class="corp-kpi-label">Pendentes</div></div><div class="corp-kpi-card" style="--kpi-color:#16a34a"><div class="corp-kpi-icon">✅</div><div class="corp-kpi-value">'+o+'</div><div class="corp-kpi-label">Concluídas</div></div><div class="corp-kpi-card" style="--kpi-color:#dc2626"><div class="corp-kpi-icon">❌</div><div class="corp-kpi-value">'+a+'</div><div class="corp-kpi-label">Canceladas</div></div>';}}function roboAtualizarContador(){if(!roboTbody)return;const e=roboTbody.querySelectorAll("tr").length;roboCounter.textContent=`${e} / 26`;const t=e>=26;roboCounter.classList.toggle("ok",t),roboBtnAdd.disabled=t,roboBtnRem.disabled=0===e}function roboCriarLinha(){const e=document.createElement("tr");return e.innerHTML='\n    <td><input type="text" /></td>\n    <td>\n      <select>\n        <option></option>\n        <option>Inserir</option>\n        <option>Retirar</option>\n        <option>Transferir</option>\n      </select>\n    </td>\n    <td><input type="text" class="robo-col-data" readonly /></td>\n    <td><input type="text" class="robo-col-hora" readonly /></td>\n    <td><input type="text" /></td>\n    <td>\n      <select class="robo-col-status" onchange="roboAtualizarDashboard()">\n        <option></option>\n        <option>Pendente</option>\n        <option>Concluída</option>\n        <option>Cancelada</option>\n      </select>\n    </td>\n    <td style="width:1%;white-space:nowrap;">\n      <div class=\"tbl-actions\">\n        <button class="btn-edit" type="button" onclick="roboToggleEditar(this)">✏️ Editar</button>\n        <button class="btn-del" type="button" onclick="roboExcluirLinha(this)">🗑️ Excluir</button>\n      </div>\n    </td>',e.querySelectorAll("input:not(.robo-col-data):not(.robo-col-hora)").forEach(e=>e.readOnly=!0),e.querySelectorAll("select").forEach(e=>e.disabled=!0),e.addEventListener("input",()=>{const t=e.querySelector(".robo-col-data"),o=e.querySelector(".robo-col-hora");if(t&&!t.value){const e=new Date;t.value=e.toLocaleDateString("pt-BR"),o.value=e.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"})}roboAtualizarDashboard()}),e}function roboAdicionarLinha(){
  if(roboTbody.querySelectorAll("tr").length>=26)return;
  const tr=roboCriarLinha();
  roboTbody.appendChild(tr);
  // Activate edit mode immediately so user can type right away
  const btn=tr.querySelector(".btn-edit");
  if(btn) roboToggleEditar(btn);
  roboAtualizarContador();
}function roboRemoverUltima(){const e=roboTbody.querySelectorAll("tr");e.length&&(roboTbody.removeChild(e[e.length-1]),roboAtualizarDashboard(),roboAtualizarContador())}function roboToggleEditar(e){const t=e.closest("tr"),isOpening=!t.classList.contains("editando");
  // close all other open rows first
  roboTbody.querySelectorAll("tr.editando").forEach(r=>{
    if(r===t)return;
    r.classList.remove("editando");
    r.querySelectorAll("input:not(.robo-col-data):not(.robo-col-hora)").forEach(i=>i.readOnly=true);
    r.querySelectorAll("select").forEach(s=>s.disabled=true);
    const btn=r.querySelector(".btn-edit");
    if(btn){btn.title="Editar linha";btn.textContent="✏️ Editar";}
  });
  const o=t.querySelectorAll("input:not(.robo-col-data):not(.robo-col-hora)"),a=t.querySelectorAll("select");
  t.classList.toggle("editando",isOpening);
  o.forEach(i=>{i.readOnly=!isOpening});a.forEach(s=>{s.disabled=!isOpening});
  t.querySelectorAll(".robo-col-data, .robo-col-hora").forEach(i=>i.readOnly=true);
  e.title=isOpening?"Confirmar edição":"Editar linha";e.textContent=isOpening?"✅":"✏️ Editar";}function roboExcluirLinha(e){if(!confirm("Excluir esta linha?"))return;e.closest("tr")?.remove(),roboAtualizarDashboard(),roboAtualizarContador()}function roboSalvar(){const e=document.getElementById("robo_ticket").value.trim(),t=document.getElementById("robo_resp").value.trim();if(!e)return void alert("Informe o número do Ticket antes de salvar!");const o=[];roboTbody.querySelectorAll("tr").forEach(e=>{const t=e.querySelectorAll("input"),a=e.querySelectorAll("select");o.push({fita:t[0].value,acao:a[0].value,data:t[1].value,hora:t[2].value,observacao:t[3].value,status:a[1].value})});const a=JSON.parse(localStorage.getItem("robo_registros")||"[]");a.unshift({ticket:e,responsavel:t,data:(new Date).toLocaleString(),linhas:o}),localStorage.setItem("robo_registros",JSON.stringify(a)),registrarLog("Robô de Fitas","Ticket salvo: "+e),alert(`Ticket ${e} salvo! (${o.length} posições)`)}function roboLoad(){
  const regs=JSON.parse(localStorage.getItem("robo_registros")||"[]");
  if(!regs.length)return void alert("Nenhum registro salvo encontrado.");
  const lista=regs.map((r,i)=>`${i+1}. Ticket ${r.ticket} — ${r.data} (${r.linhas.length} posições)`).join("\n");
  const escolha=prompt(`Registros salvos:\n\n${lista}\n\nDigite o número do registro para carregar (ou cancele):`);
  if(!escolha)return;
  const idx=parseInt(escolha)-1;
  if(isNaN(idx)||idx<0||idx>=regs.length)return void alert("Número inválido!");
  const reg=regs[idx];
  document.getElementById("robo_ticket").value=reg.ticket;
  document.getElementById("robo_resp").value=reg.responsavel||"";
  roboTbody.innerHTML="";
  reg.linhas.forEach(linha=>{
    const tr=roboCriarLinha();
    const inputs=tr.querySelectorAll("input");
    const selects=tr.querySelectorAll("select");
    // inputs[0]=fita, inputs[1]=data(robo-col-data), inputs[2]=hora(robo-col-hora), inputs[3]=obs
    if(inputs[0]) inputs[0].value=linha.fita||"";
    const dataInput=tr.querySelector(".robo-col-data");
    const horaInput=tr.querySelector(".robo-col-hora");
    if(dataInput) dataInput.value=linha.data||"";
    if(horaInput) horaInput.value=linha.hora||"";
    if(inputs[3]) inputs[3].value=linha.observacao||"";
    // selects[0]=acao, selects[1]=status
    if(selects[0]){
      // normalize saved value to match current options (Inserir/Retirar/Transferir)
      const acaoNorm={
        "ENTRADA":"Inserir","Inserir":"Inserir",
        "SAÍDA":"Retirar","Retirar":"Retirar",
        "Transferir":"Transferir"
      };
      const acaoVal=acaoNorm[linha.acao]||linha.acao||"";
      for(const opt of selects[0].options){ if(opt.value===acaoVal||opt.text===acaoVal){opt.selected=true;break;} }
    }
    if(selects[1]){
      // normalize saved value to match current options (Pendente/Concluída/Cancelada)
      const statusNorm={
        "PENDENTE":"Pendente","Pendente":"Pendente",
        "CONCLUÍDO":"Concluída","Concluída":"Concluída","Concluido":"Concluída",
        "CANCELADO":"Cancelada","Cancelada":"Cancelada"
      };
      const statusVal=statusNorm[linha.status]||linha.status||"";
      for(const opt of selects[1].options){ if(opt.value===statusVal||opt.text===statusVal){opt.selected=true;break;} }
    }
    roboTbody.appendChild(tr);
  });
  roboAtualizarDashboard();
  roboAtualizarContador();
  alert(`Registro carregado: Ticket ${reg.ticket}`);
}
function roboExportar(){if("undefined"==typeof XLSX)return void alert("Biblioteca XLSX não carregada.");const e=document.getElementById("robo_ticket").value.trim()||"sem_ticket",t=document.getElementById("robo_resp").value.trim(),o=[];if(roboTbody.querySelectorAll("tr").forEach(e=>{const t=e.querySelectorAll("input"),a=e.querySelectorAll("select");o.push({"ID Fita":t[0].value,"Ação":a[0].value,Data:t[1].value,Hora:t[2].value,"Observação":t[3].value,Status:a[1].value})}),!o.length)return void alert("Nenhuma linha para exportar!");const a=[{"ID Fita":"Ticket:","Ação":e},{"ID Fita":"Responsável:","Ação":t},{"ID Fita":"Exportado em:","Ação":(new Date).toLocaleString()},{}],n=XLSX.utils.json_to_sheet([...a,...o]),r=XLSX.utils.book_new();XLSX.utils.book_append_sheet(r,n,"Movimentação"),XLSX.writeFile(r,`movimentacao_${e}.xlsx`)}window.addEventListener("DOMContentLoaded",()=>{applyTheme(localStorage.getItem("theme")??(window.matchMedia("(prefers-color-scheme: dark)").matches?"dark":"light"));const logado=localStorage.getItem("usuarioLogado");if(logado){document.body.classList.remove("login-bg");atualizarTabela();atualizarDashboard();renderLogs();iniciarSessao(logado);}else{loginDiv.style.display="flex";menu.style.display="none";painel.style.display="none";document.body.classList.add("login-bg");atualizarTabela();atualizarDashboard();renderLogs();}}),function(){
  const STORAGE_KEY = "connmap_rows";
  const FIELDS = ["rack","device","vendor","hostname","dio","mod","pp","iface","port","direction","fl","dg","bt","sh","bl"];
  let rows = [], searchTerm = "", initialized = false;

  // Helper: busca elemento pelo id com prefixo cm_
  const el = (id) => document.getElementById("cm_" + id);
  // Helper: trim seguro
  const trim = (v) => (v ?? "").toString().trim();

  // Popula select com números de t até o
  function populateSelect(selectId, from, to) {
    const sel = el(selectId.replace("cm_",""));
    if (!sel) return;
    sel.querySelectorAll("option:not(:first-child)").forEach(o => o.remove());
    for (let num = from; num <= to; num++) {
      const opt = document.createElement("option");
      opt.value = opt.textContent = String(num);
      sel.appendChild(opt);
    }
  }

  // Limpa formulário
  function clearForm() {
    FIELDS.forEach(field => {
      const input = el(field);
      if (input) input.value = "";
    });
  }

  // Salva no localStorage
  function saveRows() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
  }

  // Carrega do localStorage
  function loadRows() {
    try {
      const stored = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
      return (Array.isArray(stored) ? stored : []).map(normalizeRow);
    } catch { return []; }
  }

  // Coleta valores do formulário
  function collectForm() {
    const getVal = (field) => trim(el(field) ? el(field).value : "");
    const obj = {
      rack:      getVal("rack"),
      device:    getVal("device"),
      vendor:    getVal("vendor"),
      hostname:  getVal("hostname"),
      dio:       getVal("dio"),
      mod:       getVal("mod"),
      pp:        getVal("pp"),
      iface:     getVal("iface"),
      port:      getVal("port"),
      direction: getVal("direction"),
      fl:        getVal("fl"),
      dg:        getVal("dg"),
      bt:        getVal("bt"),
      sh:        getVal("sh"),
      bl:        getVal("bl")
    };
    const errors = [];
    if (!obj.device) errors.push("Dispositivo é obrigatório.");
    if (obj.direction && !["A","B"].includes(obj.direction)) errors.push("Sentido deve ser A ou B.");
    return { obj, errors };
  }

  // Normaliza uma linha (lida do localStorage ou importada)
  function normalizeRow(row) {
    const direction = trim(row.direction);
    return {
      rack:       trim(row.rack),
      device:     trim(row.device),
      vendor:     trim(row.vendor),
      hostname:   trim(row.hostname),
      dio:        trim(row.dio),
      mod:        trim(row.mod),
      pp:         trim(row.pp),
      iface:      trim(row.iface),
      port:       trim(row.port),
      direction:  ["A","B"].includes(direction) ? direction : "",
      fl:         trim(row.fl),
      dg:         trim(row.dg),
      bt:         trim(row.bt),
      sh:         trim(row.sh),
      bl:         trim(row.bl),
      evidencias: Array.isArray(row.evidencias) ? row.evidencias : []
    };
  }

  // Preenche formulário com dados de uma linha (para edição)
  function fillForm(row) {
    FIELDS.forEach(field => {
      const input = el(field);
      if (input) input.value = row[field] ?? "";
    });
  }

  // Atualiza contadores KPI (elementos sem prefixo cm_)
  function updateKPI(filtered) {
    const kpis = {
      kpiTotal:     rows.length,
      kpiFiltradas: filtered.length,
      kpiA:         filtered.filter(r => r.direction === "A").length,
      kpiB:         filtered.filter(r => r.direction === "B").length
    };
    Object.entries(kpis).forEach(([id, val]) => {
      const elem = document.getElementById(id);
      if (elem) elem.textContent = val;
    });
    const rc = el("rowCount");       if (rc) rc.textContent = rows.length;
    const rcf = el("rowCountFiltered"); if (rcf) rcf.textContent = filtered.length;
  }

  // Renderiza a tabela
  function render() {
    const tbody = document.querySelector("#cm_mapTable tbody");
    if (!tbody) return;
    const filtered = searchTerm
      ? rows.filter(r => Object.values(r).join(" ").toLowerCase().includes(searchTerm.toLowerCase()))
      : rows.slice();
    updateKPI(filtered);
    tbody.innerHTML = "";
    filtered.forEach(function(rowData) {
      const realIdx = rows.indexOf(rowData);
      const tr = document.createElement("tr");
      const evs = rowData.evidencias || [];
      const evHtml = evs.length
        ? evs.slice(0,3).map((f,fi) => f.type && f.type.startsWith('image/')
            ? `<span class="ev-inline-thumb" onclick="evLightboxOpen('connmap',${realIdx},${fi})" title="${f.name}"><img src="${f.data}" alt="${f.name}"></span>`
            : `<span class="ev-inline-thumb" onclick="evLightboxOpen('connmap',${realIdx},${fi})" title="${f.name}">📄 ${f.name.substring(0,10)}…</span>`
          ).join(' ') + (evs.length > 3 ? ` <span style="font-size:11px;color:var(--sub)">+${evs.length-3}</span>` : '')
        : '<span style="font-size:11px;color:var(--sub)">—</span>';
      tr.innerHTML = `
        <td>${rowData.rack ?? "—"}</td>
        <td class="truncate t-xs" title="${rowData.device ?? ""}">${rowData.device ?? "—"}</td>
        <td class="truncate t-xs" title="${rowData.vendor ?? ""}">${rowData.vendor ?? "—"}</td>
        <td class="truncate"      title="${rowData.hostname ?? ""}">${rowData.hostname ?? "—"}</td>
        <td>${rowData.dio ?? "—"}</td><td>${rowData.mod ?? "—"}</td><td>${rowData.pp ?? "—"}</td>
        <td class="truncate t-sm" title="${rowData.iface ?? ""}">${rowData.iface ?? "—"}</td>
        <td>${rowData.port ?? "—"}</td>
        <td><span class="pill ${rowData.direction ? "ok" : "warn"}">${rowData.direction || "—"}</span></td>
        <td>${rowData.fl ?? "—"}</td><td>${rowData.dg ?? "—"}</td><td>${rowData.bt ?? "—"}</td>
        <td>${rowData.sh ?? "—"}</td><td>${rowData.bl ?? "—"}</td>
        <td>${evHtml}</td>
        <td class="actions">
          <div class="tbl-actions">
            <button class="btn-dark" data-act="ev"   title="Evidências" type="button">📎</button>
            <button class="btn-edit" data-act="edit" title="Editar"      type="button">✏️ Editar</button>
            <button class="btn-del"  data-act="del"  title="Excluir"     type="button">🗑️ Excluir</button>
          </div>
        </td>`;
      tr.querySelectorAll("button").forEach(btn => {
        btn.addEventListener("click", () => {
          const act = btn.getAttribute("data-act");
          if (act === "del") {
            if (confirm("Excluir esta linha?")) {
              rows.splice(realIdx, 1);
              saveRows();
              render();
            }
          } else if (act === "edit") {
            fillForm(rows[realIdx]);
            rows.splice(realIdx, 1);
            saveRows();
            render();
          } else if (act === "ev") {
            abrirEvidencias('connmap', realIdx);
          }
        });
      });
      tr.addEventListener("contextmenu", e => {
        e.preventDefault();
        const choice = prompt("Ação:\n1 = Inserir acima\n2 = Inserir abaixo\n3 = Excluir");
        if (!choice) return;
        if (choice === "1") {
          const { obj } = collectForm();
          rows.splice(realIdx, 0, normalizeRow(obj));
          saveRows(); render();
        } else if (choice === "2") {
          const { obj } = collectForm();
          rows.splice(realIdx + 1, 0, normalizeRow(obj));
          saveRows(); render();
        } else if (choice === "3") {
          if (confirm("Excluir esta linha?")) {
            rows.splice(realIdx, 1);
            saveRows(); render();
          }
        }
      });
      tbody.appendChild(tr);
    });
  }

  // Gera HTML para exportação
  function buildExportHTML() {
    const headers = ["Rack","Dispositivo","Fabricante","Hostname","DIO","MOD","PP","Interface","Porta","Sentido","FL","DG","BT","SH","BL"];
    const rowsHtml = rows.map(r =>
      "<tr>" + [r.rack,r.device,r.vendor,r.hostname,r.dio,r.mod,r.pp,r.iface,r.port,r.direction,r.fl,r.dg,r.bt,r.sh,r.bl]
        .map(v => `<td>${String(v ?? "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;") || ""}</td>`)
        .join("") + "</tr>"
    ).join("\n");
    return `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>
      <table border="1"><thead><tr>${headers.map(h=>`<th>${h}</th>`).join("")}</tr></thead>
      <tbody>${rowsHtml}</tbody></table></body></html>`;
  }

  // Exporta para Excel
  function exportExcel() {
    const blob = new Blob([buildExportHTML()], { type: "application/vnd.ms-excel" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "connmap.xls";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  // Exporta PDF
  function exportPdf() {
    const win = window.open("", "_blank");
    const tableHtml = buildExportHTML().replace(/^[\s\S]*<table/,"<table").replace(/<\/table>[\s\S]*$/,"</table>");
    win.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>ConnMap - Export PDF</title>
      <style>body{font-family:Segoe UI,Arial,sans-serif;padding:16px}h1{font-size:18px;margin:0 0 12px}
      table{width:100%;border-collapse:collapse;font-size:12px}th,td{border:1px solid #999;padding:6px 8px}
      th{background:#f0f0f0}</style></head><body><h1>ConnMap (De/Para)</h1>${tableHtml}</body></html>`);
    win.document.close();
    win.focus();
    win.print();
  }

  // Parse CSV simples
  function parseCsvLine(line) {
    const result = []; let cur = "", inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"' && line[i+1] === '"' && inQ) { cur += '"'; i++; }
      else if (ch === '"') { inQ = !inQ; }
      else if (ch === ',' && !inQ) { result.push(cur); cur = ""; }
      else { cur += ch; }
    }
    result.push(cur);
    return result.map(v => v.replace(/^"(.*)"$/, "$1").replace(/""/g, '"'));
  }

  function parseCsv(text) {
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (!lines.length) return [];
    const headers = parseCsvLine(lines.shift());
    const idx = (h) => headers.indexOf(h);
    return lines.map(line => {
      const cols = parseCsvLine(line);
      const get = (h) => idx(h) >= 0 ? cols[idx(h)] ?? "" : "";
      return {
        rack: get("rack"), device: get("device"), vendor: get("vendor"), hostname: get("hostname"),
        dio: get("dio"), mod: get("mod"), pp: get("pp"), iface: get("iface"), port: get("port"),
        direction: get("direction"), fl: get("fl"), dg: get("dg"), bt: get("bt"), sh: get("sh"), bl: get("bl"),
        _legacy_mod: get("MOD-BL-PP") || get("mod_bl_pp"),
        _legacy_fl:  get("FL-DIO-DG") || get("fl_dio_dg"),
        _legacy_bt:  get("BT-SH")     || get("bt_sh")
      };
    });
  }

  window.cmInit = function() {
    if (initialized) return;
    initialized = true;

    populateSelect("dio", 1, 45);
    populateSelect("mod", 1, 4);

    // Botão Adicionar
    const addBtn = el("addBtn");
    if (addBtn) {
      addBtn.addEventListener("click", () => {
        const { obj, errors } = collectForm();
        if (errors.length) { alert(errors.join("\n")); return; }
        rows.push(normalizeRow(obj));
        clearForm();
        saveRows();
        render();
      });
    }

    // Botão Limpar formulário
    const resetBtn = el("resetBtn");
    if (resetBtn) resetBtn.addEventListener("click", clearForm);

    // Exportar Excel
    const excelBtn = el("exportExcelBtn");
    if (excelBtn) excelBtn.addEventListener("click", exportExcel);

    // Exportar PDF
    const pdfBtn = el("exportPdfBtn");
    if (pdfBtn) pdfBtn.addEventListener("click", exportPdf);

    // Importar arquivo
    const importFile = el("importFile");
    if (importFile) {
      importFile.addEventListener("change", e => {
        const file = e.target.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = evt => {
          try {
            const text = evt.target.result;
            let imported = file.name.toLowerCase().endsWith(".json")
              ? JSON.parse(text)
              : parseCsv(text);
            if (!Array.isArray(imported)) throw new Error("Formato inválido");
            imported = imported.map(normalizeRow);
            rows = rows.concat(imported);
            saveRows();
            render();
            alert("Importadas " + imported.length + " linhas.");
          } catch (err) {
            alert("Falha ao importar: " + err.message);
          } finally {
            e.target.value = "";
          }
        };
        reader.readAsText(file);
      });
    }

    // Pesquisa
    const searchBox = el("searchBox");
    if (searchBox) {
      searchBox.addEventListener("input", e => { searchTerm = e.target.value.trim(); render(); });
      searchBox.addEventListener("keydown", e => {
        if (e.key === "Escape") { e.target.value = ""; searchTerm = ""; render(); }
      });
    }

    // Expor render para uso externo (ex: evidências)
    window.cmRender = render;

    // Carregar dados salvos e renderizar
    rows = loadRows();
    render();
  };
}()
const DSC_KEY="descarte";function dscGetData(){try{return JSON.parse(localStorage.getItem(DSC_KEY)||"[]")}catch(e){return[]}}function dscSaveData(e){localStorage.setItem(DSC_KEY,JSON.stringify(e))}!function(){function e(e){const t=document.getElementById("dsc_tbody");if(!t)return;const o=dscGetData();t.innerHTML=e.map((e,t)=>{const a=o.indexOf(e),n=a>=0?a:t;let r="";const i=e.evidencias||[];return i.length?(r=i.slice(0,3).map((e,t)=>e.type&&e.type.startsWith("image/")?`<span class="ev-inline-thumb" onclick="evLightboxOpen(${n},${t})" title="${esc(e.name)}"><img src="${e.data}" alt="${esc(e.name)}"></span>`:`<span class="ev-inline-thumb" onclick="evLightboxOpen(${n},${t})" title="${esc(e.name)}">📄 ${esc(e.name.substring(0,10))}…</span>`).join(" "),i.length>3&&(r+=` <span style="font-size:11px;color:var(--sub)">+${i.length-3}</span>`)):r='<span style="font-size:11px;color:var(--sub)">—</span>',`\n      <tr>\n        <td>${t+1}</td>\n        <td>${e.data||""}</td>\n        <td>${e.solicitante||""}</td>\n        <td>${e.descricao||""}</td>\n        <td>${e.marca||""}</td>\n        <td>${e.modelo||""}</td>\n        <td>${e.serial||""}</td>\n        <td>${e.sgp||""}</td>\n        <td>${e.motivo||""}</td>\n        <td>${e.extras||""}</td>\n        <td>${r}</td>\n        <td style="display:flex;gap:6px;justify-content:center;align-items:center">\n          <div class="tbl-actions"><button class="btn-dark" onclick="dscAbrirEvidencia(${n})" type="button" title="Evidências">📎</button><button class="btn-edit" onclick="dscEditar(${n})" type="button">✏️ Editar</button><button class="btn-del" onclick="dscDel(${n})" type="button">🗑️ Excluir</button></div>\n        </td>\n      </tr>`}).join("")||'<tr><td colspan="12" style="padding:16px; text-align:center; color:var(--sub)">Nenhum registro encontrado.</td></tr>'}window.dscAddCampo=function(){const e=document.getElementById("dsc_extras"),t=document.createElement("div");t.style.cssText="display:flex; align-items:center; gap:8px; margin-top:8px",t.innerHTML='<input type="text" class="dsc_extra_input" placeholder="Outro ID (MAC, IMEI...)" style="flex:1">\n      <button type="button" onclick="this.parentElement.remove()" style="background:none; border:none; color:var(--bad); font-size:20px; cursor:pointer; padding:0 4px; min-height:unset">&times;</button>',e.appendChild(t)},window.dscLimpar=function(){dscEditIndex=null;["dsc_solicitante","dsc_centro","dsc_local","dsc_motivo","dsc_descricao","dsc_marca","dsc_modelo","dsc_serial","dsc_partnumber","dsc_sgp"].forEach(e=>{document.getElementById(e).value=""}),document.getElementById("dsc_extras").innerHTML="",document.getElementById("dsc_tipo").value="Usado"},window.dscToggleForm=function(){const e=document.getElementById("formDescarteWrapper"),t=document.getElementById("btnToggleDescarte"),o="none"!==e.style.display;e.style.display=o?"none":"block",t.textContent=o?"➕ Novo Registro":"➖ Fechar Formulário"},window.dscFiltrar=function(t){const o=t.toLowerCase(),a=dscGetData();e(o?a.filter(e=>(e.solicitante||"").toLowerCase().includes(o)||(e.descricao||"").toLowerCase().includes(o)||(e.serial||"").toLowerCase().includes(o)||(e.sgp||"").toLowerCase().includes(o)||(e.motivo||"").toLowerCase().includes(o)):a)},window.descarteRender=function(){e(dscGetData())},window.dscDel=function(e){if(!confirm("Excluir este registro?"))return;const t=dscGetData();t.splice(e,1),dscSaveData(t),registrarLog("Descarte","Registro removido"),descarteRender()},window.dscEditar=function(e){const t=dscGetData()[e];if(!t)return;dscEditIndex=e;const o=document.getElementById("formDescarteWrapper"),a=document.getElementById("btnToggleDescarte");o.style.display="block",a.textContent="➖ Fechar Formulário",document.getElementById("dsc_solicitante").value=t.solicitante||"",document.getElementById("dsc_centro").value=t.centro||"",document.getElementById("dsc_local").value=t.local||"",document.getElementById("dsc_motivo").value=t.motivo||"",document.getElementById("dsc_descricao").value=t.descricao||"",document.getElementById("dsc_tipo").value=t.tipo||"Usado",document.getElementById("dsc_marca").value=t.marca||"",document.getElementById("dsc_modelo").value=t.modelo||"",document.getElementById("dsc_serial").value=t.serial||"",document.getElementById("dsc_partnumber").value=t.partnumber||"",document.getElementById("dsc_sgp").value=t.sgp||"";const n=document.getElementById("dsc_extras");n.innerHTML="",t.extras&&t.extras.split(" | ").forEach(e=>{const t=document.createElement("div");t.style.cssText="display:flex; align-items:center; gap:8px; margin-top:8px",t.innerHTML=`<input type="text" class="dsc_extra_input" value="${e}" style="flex:1">\n          <button type="button" onclick="this.parentElement.remove()" style="background:none; border:none; color:var(--bad); font-size:20px; cursor:pointer; padding:0 4px; min-height:unset">&times;</button>`,n.appendChild(t)});const r=document.querySelector('[onclick="dscRegistrar()"]');r&&(r.textContent="💾 Salvar Alterações",r.setAttribute("onclick",`dscSalvarEdicao(${e})`)),o.scrollIntoView({behavior:"smooth",block:"start"})},window.dscSalvarEdicao=function(e){const t=document.getElementById("dsc_solicitante").value.trim(),o=document.getElementById("dsc_descricao").value.trim(),a=document.getElementById("dsc_serial").value.trim();if(!t||!o)return void alert("Preencha ao menos Solicitante e Descrição do Item.");const n=[];document.querySelectorAll(".dsc_extra_input").forEach(e=>{e.value&&n.push(e.value)});const r=dscGetData();r[e]={data:r[e].data,solicitante:t,centro:document.getElementById("dsc_centro").value,tipo:document.getElementById("dsc_tipo").value,local:document.getElementById("dsc_local").value,motivo:document.getElementById("dsc_motivo").value,descricao:o,marca:document.getElementById("dsc_marca").value,modelo:document.getElementById("dsc_modelo").value,serial:a,partnumber:document.getElementById("dsc_partnumber").value,sgp:document.getElementById("dsc_sgp").value,extras:n.join(" | "),evidencias:r[e].evidencias||[]},dscSaveData(r),registrarLog("Descarte","Registro editado: "+o),alert("Registro atualizado com sucesso!"),dscLimpar(),descarteRender();const i=document.querySelector('[onclick^="dscSalvarEdicao"]');i&&(i.textContent="✅ Registrar Equipamento",i.setAttribute("onclick","dscRegistrar()"))},window.dscRegistrar=function(){const e=document.getElementById("dsc_solicitante").value.trim(),t=document.getElementById("dsc_descricao").value.trim(),o=document.getElementById("dsc_serial").value.trim();if(!e||!t)return void alert("Preencha ao menos Solicitante e Descrição do Item.");const a=[];document.querySelectorAll(".dsc_extra_input").forEach(e=>{e.value&&a.push(e.value)});const n={data:(new Date).toLocaleString("pt-BR"),solicitante:e,centro:document.getElementById("dsc_centro").value,tipo:document.getElementById("dsc_tipo").value,local:document.getElementById("dsc_local").value,motivo:document.getElementById("dsc_motivo").value,descricao:t,marca:document.getElementById("dsc_marca").value,modelo:document.getElementById("dsc_modelo").value,serial:o,partnumber:document.getElementById("dsc_partnumber").value,sgp:document.getElementById("dsc_sgp").value,extras:a.join(" | "),evidencias:[]},r=dscGetData();null!==dscEditIndex?(r[dscEditIndex]=n,dscEditIndex=null,registrarLog("Descarte","Equipamento editado: "+t),alert("Registro atualizado com sucesso!")):(r.push(n),registrarLog("Descarte","Equipamento registrado: "+t),alert("Equipamento registrado com sucesso!")),dscSaveData(r),dscLimpar(),descarteRender()},window.dscExportarExcel=function(){const e=dscGetData();if(!e.length)return void alert("Nenhum registro para exportar.");const t=XLSX.utils.json_to_sheet(e.map(e=>({Data:e.data,Solicitante:e.solicitante,Centro:e.centro,Tipo:e.tipo,"Localização":e.local,Motivo:e.motivo,"Descrição":e.descricao,Marca:e.marca,Modelo:e.modelo,"Serial (S/N)":e.serial,"Part Number":e.partnumber,SGP:e.sgp,Extras:e.extras}))),o=XLSX.utils.book_new();XLSX.utils.book_append_sheet(o,t,"Descarte"),XLSX.writeFile(o,"inventario_descarte.xlsx")}}(),function(){
  let _evIdx=null, _evFiles=[], _evCtx=null;

  // Contextos: descarte e desembalagem (entrada/saida/coleta)
  const _evSources = {
    descarte:  { getData: dscGetData,      saveData: dscSaveData,      labelFn: (r,i) => r.descricao||'Registro #'+(i+1),     renderFn: () => { if(typeof descarteRender==='function') descarteRender(); } },
    desemb_entrada: { getData: getDesembEntrada, saveData: saveDesembEntrada, labelFn: (r,i) => 'Entrada — Caixa: '+(r.caixa||'#'+(i+1)),  renderFn: desembRenderEntrada },
    desemb_saida:   { getData: getDesembSaida,   saveData: saveDesembSaida,   labelFn: (r,i) => 'Saída — Caixa: '+(r.caixa||'#'+(i+1)),    renderFn: desembRenderSaida   },
    desemb_coleta:  { getData: getDesembColeta,  saveData: saveDesembColeta,  labelFn: (r,i) => 'Coleta — Caixa: '+(r.caixa||'#'+(i+1)),   renderFn: desembRenderColeta  },
    connmap:        { getData: () => JSON.parse(localStorage.getItem('connmap_rows')||'[]'),
                      saveData: d  => localStorage.setItem('connmap_rows', JSON.stringify(d)),
                      labelFn: (r,i) => 'ConnMap #'+(i+1)+(r.device?' — '+r.device:''),
                      renderFn: () => { if(typeof window.cmRender==='function') window.cmRender(); } },
  };

  function _evRenderPreview() {
    const grid = document.getElementById('evPreviewGrid');
    if(!grid) return;
    grid.innerHTML = _evFiles.map((f,i) => `<div class="ev-thumb">
      ${f.type&&f.type.startsWith('image/')?`<img src="${f.data}" alt="${esc(f.name)}">`:'<div class="ev-pdf-icon">📄</div>'}
      <button class="ev-rm" onclick="_evRemover(${i})" title="Remover" type="button">✕</button>
    </div>`).join('');
  }

  // Abrir evidências — função genérica
  window.abrirEvidencias = function(source, idx) {
    const ctx = _evSources[source];
    if(!ctx) return;
    const data = ctx.getData();
    if(idx < 0 || idx >= data.length) return alert('Registro não encontrado.');
    _evCtx = { source, idx, ctx };
    _evIdx = idx;
    _evFiles = JSON.parse(JSON.stringify(data[idx].evidencias || []));
    _evRenderPreview();
    document.getElementById('modalEvidencias').classList.add('open');
    document.querySelector('#modalEvidencias h3').textContent = '📎 Evidências — ' + ctx.labelFn(data[idx], idx);
  };

  // Manter compatibilidade com descarte (chamadas antigas)
  window.dscAbrirEvidencia = function(idx) { abrirEvidencias('descarte', idx); };

  window.dscAbrirEvidenciaGlobal = function(){
    const data = dscGetData();
    if(!data.length) return alert('Nenhum registro de descarte encontrado. Registre um item primeiro.');
    document.getElementById('evSelectList').innerHTML = data.map((r,i) => `
      <div class="ev-select-item" onclick="evSelecionarRegistro(${i})">
        <strong>#${i+1} — ${esc(r.descricao||'Sem descrição')}</strong>
        <span>${esc(r.serial||'s/serial')} ${r.sgp?'· SGP: '+esc(r.sgp):''}</span>
        ${(r.evidencias||[]).length?`<em>${r.evidencias.length} evidência(s) já vinculada(s)</em>`:''}
      </div>`).join('');
    document.getElementById('modalEvSelect').classList.add('open');
  };
  window.evSelecionarRegistro = function(idx) {
    document.getElementById('modalEvSelect').classList.remove('open');
    abrirEvidencias('descarte', idx);
  };

  // Drop zone
  const dz = document.getElementById('evDropZone');
  if(dz) {
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('dragover'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('dragover'); evHandleFiles(e.dataTransfer.files); });
  }

  window.evHandleFiles = function(files) {
    Array.from(files).forEach(f => {
      const r = new FileReader();
      r.onload = ev => { _evFiles.push({ name: f.name, type: f.type, data: ev.target.result }); _evRenderPreview(); };
      r.readAsDataURL(f);
    });
  };
  window._evRemover = function(i) { _evFiles.splice(i,1); _evRenderPreview(); };

  window.evFechar = function() {
    document.getElementById('modalEvidencias').classList.remove('open');
    document.getElementById('evFileInput').value = '';
    _evFiles = []; _evIdx = null; _evCtx = null;
  };

  window.evSalvar = function() {
    if(!_evCtx) return;
    const savedCtx = _evCtx; // salva referência antes de evFechar() limpar _evCtx
    const data = savedCtx.ctx.getData();
    if(!data[savedCtx.idx]) return;
    data[savedCtx.idx].evidencias = _evFiles;
    savedCtx.ctx.saveData(data);
    evFechar();
    savedCtx.ctx.renderFn && savedCtx.ctx.renderFn();
    alert('Evidências salvas com sucesso!');
  };

  // Lightbox
  let _lbIdx = null, _lbFiles = [];
  function _lbShow() {
    const f = _lbFiles[_lbIdx]; if(!f) return;
    const img  = document.getElementById('evLightboxImg');
    const pdf  = document.getElementById('evLightboxPdf');
    const link = document.getElementById('evLightboxLink');
    const name = document.getElementById('evLightboxName');
    if(f.type&&f.type.startsWith('image/')) { img.src=f.data; img.style.display='block'; pdf.style.display='none'; }
    else { img.style.display='none'; pdf.style.display='block'; name.textContent=f.name; link.href=f.data; }
  }
  window.evLightboxOpen = function(source, rowIdx, fileIdx) {
    // suporte para chamada antiga evLightboxOpen(rowIdx, fileIdx) do descarte
    if(typeof source === 'number') { fileIdx=rowIdx; rowIdx=source; source='descarte'; }
    const ctx = _evSources[source];
    if(!ctx) return;
    const data = ctx.getData();
    _lbFiles = (data[rowIdx] ? data[rowIdx].evidencias : null) || [];
    _lbIdx = fileIdx; _lbShow();
    document.getElementById('evLightbox').classList.add('open');
  };
  window.evLightboxNav = function(d) { _lbIdx = (_lbIdx+d+_lbFiles.length)%_lbFiles.length; _lbShow(); };
  window.evLightboxClose = function() { document.getElementById('evLightbox').classList.remove('open'); };

  document.getElementById('evLightbox').addEventListener('click', function(e){ if(e.target===this) evLightboxClose(); });
  document.getElementById('modalEvidencias').addEventListener('click', function(e){ if(e.target===this) evFechar(); });
  document.getElementById('modalEvSelect').addEventListener('click', function(e){ if(e.target===this) this.classList.remove('open'); });
}()
// ===== CORPORATE DASHBOARD =====
window.abrirDashboardCorp = function() {
  document.getElementById('menu').style.display = 'none';
  document.getElementById('painel').style.display = 'block';
  mostrarAba('corpDashboard');
  const sub = document.getElementById('painelSubtitle');
  if (sub) sub.textContent = 'Dashboard Corporativo';
};

let _corpCharts = {};

function renderCorpDashboard() {
  function safeJSON(key, fb) {
    try { return JSON.parse(localStorage.getItem(key) || 'null') || fb; } catch(e) { return fb; }
  }
  function countBy(arr, key) {
    return arr.reduce((acc, item) => { const v = item[key]||'Outros'; acc[v]=(acc[v]||0)+1; return acc; }, {});
  }
  function setBadge(id, text) { const el=document.getElementById(id); if(el) el.textContent=text; }
  function kpiCard(icon, value, label, sub, color) {
    return `<div class="corp-kpi-card" style="--kpi-color:${color}">
      <div class="corp-kpi-row">
        <div class="corp-kpi-icon" style="background:color-mix(in srgb,${color} 12%,transparent)">${icon}</div>
      </div>
      <div class="corp-kpi-value">${value}</div>
      <div class="corp-kpi-label">${label}</div>
      <div class="corp-kpi-sub"><span class="kpi-dot"></span>${sub}</div>
    </div>`;
  }

  // ── Paleta definida pelo usuário ─────────────────────────────────────────
  const C = {
    blue    : '#1F3A5F',   // azul escuro — total / referência
    orange  : '#F28E2B',   // laranja — alerta / pendente / destaque
    green   : '#59A14F',   // verde — positivo / ok / concluído
    red     : '#E15759',   // rosa forte — negativo / cancelado / fora de garantia
    // aliases removidos (redLight, amber, teal, purple, brown) — use os nomes canônicos acima
    amber   : '#F28E2B',   // alias de orange (mantido por compatibilidade interna)
    teal    : '#1F3A5F',   // alias de blue (mantido por compatibilidade interna)
    purple  : '#E15759',   // alias de red (mantido por compatibilidade interna)
    slate   : '#6B7C93',   // azul acinzentado — neutro
    steel   : '#8FA3B1',   // neutro secundário
    silver  : '#C5CDD5',   // indefinido / outros
  };

  // Paleta sequencial para gráficos multi-categoria (fornecedor, local, vendor)
  const SEQ = [C.blue, C.orange, C.green, C.red, '#2E6DA4', '#D4843E', '#3D7A35', '#C03F41', C.slate, C.steel];

  // ── Load data ──
  const inv      = safeJSON('inventario', []);
  const dsc      = safeJSON('descarte', null) || safeJSON('descarte_registros', null) || safeJSON('dsc_registros', []);
  const fitas    = safeJSON('fitas', []);
  const roboRegs = safeJSON('robo_registros', []);
  const desEnt   = safeJSON('desemb_entrada', []);
  const desSai   = safeJSON('desemb_saida', []);
  const desCol   = safeJSON('desemb_coleta', []);
  const ops      = safeJSON('operacoes', []);
  const connMap  = safeJSON('connmap_rows', []);
  const opPdfs   = safeJSON('operacoes_pdfs', []);

  const anyReal = inv.length||dsc.length||fitas.length||roboRegs.length||desEnt.length||desSai.length||desCol.length||ops.length;
  const badge = document.getElementById('corpMockBadge');
  if(badge) badge.style.display = anyReal ? 'none' : 'inline-block';

  const servers = inv.length ? inv : [
    {fornecedor:'Dell',   local:'SP-DC01', status:'✅ Em Garantia',  eol:'2026-01', eos:''},
    {fornecedor:'Dell',   local:'SP-DC01', status:'✅ Em Garantia',  eol:'',        eos:''},
    {fornecedor:'HP',     local:'SP-DC01', status:'❌ Fora Garantia',eol:'2024-03', eos:'2024-01'},
    {fornecedor:'HP',     local:'RJ-DC02', status:'✅ Em Garantia',  eol:'',        eos:''},
    {fornecedor:'IBM',    local:'RJ-DC02', status:'❌ Fora Garantia',eol:'2023-12', eos:'2023-06'},
    {fornecedor:'Lenovo', local:'MG-DC03', status:'✅ Em Garantia',  eol:'',        eos:''},
    {fornecedor:'Cisco',  local:'SP-DC01', status:'❌ Fora Garantia',eol:'2024-06', eos:''},
    {fornecedor:'Dell',   local:'MG-DC03', status:'✅ Em Garantia',  eol:'',        eos:''},
    {fornecedor:'HP',     local:'SP-DC01', status:'✅ Em Garantia',  eol:'2027-01', eos:''},
  ];
  const descartes = dsc.length ? dsc : [{tipo:'Usado'},{tipo:'Defeito'},{tipo:'Obsoleto'},{tipo:'Usado'},{tipo:'Defeito'}];

  // ── Métricas ──
  const srvTotal = servers.length;
  const srvGar   = servers.filter(s=>(s.status||'').includes('Em Garantia')).length;
  const srvFora  = servers.filter(s=>(s.status||'').includes('Fora Garantia')).length;
  const srvEol   = servers.filter(s=>s.eol&&s.eol.trim()).length;
  const srvEos   = servers.filter(s=>s.eos&&s.eos.trim()).length;
  const garPct   = srvTotal ? Math.round(srvGar/srvTotal*100) : 0;

  let roboTotal=0, roboPend=0, roboConcl=0, roboCanc=0;
  roboRegs.forEach(reg=>{
    (reg.linhas||[]).forEach(l=>{
      const st=(l.status||'').toUpperCase();
      roboTotal++;
      if(st==='PENDENTE') roboPend++;
      if(st==='CONCLUÍDA'||st==='CONCLUÍDO') roboConcl++;
      if(st==='CANCELADA'||st==='CANCELADO') roboCanc++;
    });
  });
  const desembTotal = desEnt.length+desSai.length+desCol.length;
  const fitasAtivas = fitas.filter(f=>(f.status||'').toLowerCase()!=='inativo').length;

  const cmTotal   = connMap.length;
  const cmA       = connMap.filter(r=>r.direction==='A').length;
  const cmB       = connMap.filter(r=>r.direction==='B').length;
  const cmVendors = [...new Set(connMap.map(r=>r.vendor).filter(Boolean))].length;
  const cmRacks   = [...new Set(connMap.map(r=>r.rack).filter(Boolean))].length;

  // ── Chart globals ──
  const isDark    = document.body.classList.contains('dark');
  const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.05)';
  const tickColor = isDark ? '#9ca3af' : '#6b7280';
  const chartDefaults = {
    responsive:true, maintainAspectRatio:false,
    plugins:{ legend:{display:false}, tooltip:{backgroundColor:isDark?'#1f2937':'#fff',titleColor:isDark?'#f3f4f6':'#111827',bodyColor:isDark?'#d1d5db':'#374151',borderColor:isDark?'#374151':'#e5e7eb',borderWidth:1,padding:10,cornerRadius:8} }
  };
  const scalesXY = {
    x:{ticks:{color:tickColor,font:{size:11}},grid:{color:gridColor}},
    y:{ticks:{color:tickColor,font:{size:11},stepSize:1},grid:{color:gridColor},beginAtZero:true}
  };
  function destroyChart(id) {
    if(_corpCharts[id]){ try{_corpCharts[id].destroy();}catch(e){} delete _corpCharts[id]; }
  }
  function legendOpts(pos='bottom') {
    return { display:true, position:pos, labels:{color:tickColor,padding:14,font:{size:11},boxWidth:12,boxHeight:12,usePointStyle:true,pointStyle:'rectRounded'} };
  }

  // ════════════════════════════════════════════
  // MÓDULO 1 — INVENTÁRIO
  // ════════════════════════════════════════════
  const statsEl = document.getElementById('corpDashStats');
  if(statsEl) statsEl.innerHTML = [
    kpiCard('🖥️', srvTotal, 'Total de Servidores',  inv.length?'inventário real':'dados de exemplo', C.blue),
    kpiCard('✅', srvGar,   'Em Garantia',            garPct+'% do parque',                           C.green),
    kpiCard('⚠️', srvFora,  'Fora de Garantia',       srvTotal?(100-garPct)+'% do parque':'—',        C.amber),
    kpiCard('📅', srvEol,   'Com EOL Definido',       'fim de ciclo de vida',                         C.red),
    kpiCard('📋', srvEos,   'Com EOS Definido',       'fim de suporte',                               C.slate),
    kpiCard('🏭', Object.keys(countBy(servers,'fornecedor')).length, 'Fornecedores', 'no inventário', C.purple),
  ].join('');

  setBadge('badgeFornecedor', Object.keys(countBy(servers,'fornecedor')).length+' fornecedores');
  setBadge('badgeGarantia',   garPct+'% em garantia');
  setBadge('badgeLocal',      Object.keys(countBy(servers,'local')).length+' localidades');

  const garEl = document.getElementById('corpGarProgresso');
  if(garEl) garEl.innerHTML = `
    <div class="corp-prog-label"><span>Em Garantia</span><span>${srvGar} / ${srvTotal}</span></div>
    <div class="corp-prog-track"><div class="corp-prog-fill" style="width:${garPct}%;background:${C.green}"></div></div>
    <div class="corp-prog-label" style="margin-top:8px"><span>Fora de Garantia</span><span>${srvFora} / ${srvTotal}</span></div>
    <div class="corp-prog-track"><div class="corp-prog-fill" style="width:${srvTotal?Math.round(srvFora/srvTotal*100):0}%;background:${C.red}"></div></div>`;

  // Chart 1 — Fornecedor (cada barra cor própria da SEQ)
  destroyChart('fornecedor');
  const fornMap    = countBy(servers,'fornecedor');
  const fornLabels = Object.keys(fornMap).sort((a,b)=>fornMap[b]-fornMap[a]);
  const c1 = document.getElementById('chartFornecedor');
  if(c1) _corpCharts['fornecedor'] = new Chart(c1, {
    type:'bar',
    data:{ labels:fornLabels, datasets:[{ data:fornLabels.map(k=>fornMap[k]), backgroundColor:fornLabels.map((_,i)=>SEQ[i%SEQ.length]), borderRadius:6, borderSkipped:false }] },
    options:{...chartDefaults, scales:scalesXY}
  });

  // Chart 2 — Garantia (verde / vermelho / âmbar)
  destroyChart('garantia');
  const c2 = document.getElementById('chartGarantia');
  if(c2) _corpCharts['garantia'] = new Chart(c2, {
    type:'doughnut',
    data:{ labels:['Em Garantia','Fora de Garantia','Com EOL'], datasets:[{ data:[srvGar,srvFora,srvEol], backgroundColor:[C.green, C.red, C.amber], borderWidth:0, hoverOffset:8 }] },
    options:{...chartDefaults, cutout:'68%', plugins:{...chartDefaults.plugins, legend:legendOpts()}}
  });

  // Chart 3 — Local (cada barra cor própria)
  destroyChart('local');
  const localMap    = countBy(servers,'local');
  const localLabels = Object.keys(localMap).sort((a,b)=>localMap[b]-localMap[a]).slice(0,8);
  const c3 = document.getElementById('chartLocal');
  if(c3) _corpCharts['local'] = new Chart(c3, {
    type:'bar',
    data:{ labels:localLabels, datasets:[{ data:localLabels.map(k=>localMap[k]), backgroundColor:localLabels.map((_,i)=>SEQ[i%SEQ.length]), borderRadius:6, borderSkipped:false }] },
    options:{...chartDefaults, indexAxis:'y', scales:{ x:{ticks:{color:tickColor,font:{size:11},stepSize:1},grid:{color:gridColor},beginAtZero:true}, y:{ticks:{color:tickColor,font:{size:11}},grid:{color:gridColor}} }}
  });

  // ════════════════════════════════════════════
  // MÓDULO 2 — ROBÔ DE FITAS
  // ════════════════════════════════════════════
  const roboEl = document.getElementById('corpRoboKpis');
  if(roboEl) roboEl.innerHTML = [
    kpiCard('🤖', roboTotal,    'Total Movimentações', roboRegs.length+' registros salvos', C.blue),
    kpiCard('⏳', roboPend,     'Pendentes',            'aguardando execução',               C.amber),
    kpiCard('✅', roboConcl,    'Concluídas',           'finalizadas com sucesso',            C.green),
    kpiCard('❌', roboCanc,     'Canceladas',           'não executadas',                     C.red),
    kpiCard('📼', fitas.length, 'Fitas Cadastradas',    fitasAtivas+' ativas',               C.teal),
    kpiCard('📦', desembTotal,  'Desembalagem',          'entradas, saídas e coletas',         C.purple),
  ].join('');

  setBadge('badgeFitas',    fitas.length+' fitas');
  setBadge('badgeDescarte', descartes.length+' registros');
  setBadge('badgeDesemb',   desembTotal+' movimentações');

  // Chart 4 — Fitas por Status (cor semântica por status)
  destroyChart('fitas');
  const fitasRaw    = fitas.length ? fitas : [{status:'Ativo'},{status:'Ativo'},{status:'Em Uso'},{status:'Inativo'}];
  const fitasMap    = countBy(fitasRaw,'status');
  const fitasLabels = Object.keys(fitasMap);
  const fitaColor   = s => {
    const sl=(s||'').toLowerCase();
    if(sl.includes('ativo')&&!sl.includes('in')) return C.green;
    if(sl.includes('uso'))    return C.teal;
    if(sl.includes('inati'))  return C.slate;
    if(sl.includes('danif')||sl.includes('defeito')) return C.red;
    return SEQ[fitasLabels.indexOf(s)%SEQ.length];
  };
  const c4 = document.getElementById('chartFitas');
  if(c4) _corpCharts['fitas'] = new Chart(c4, {
    type:'doughnut',
    data:{ labels:fitasLabels, datasets:[{ data:fitasLabels.map(k=>fitasMap[k]), backgroundColor:fitasLabels.map(fitaColor), borderWidth:0, hoverOffset:8 }] },
    options:{...chartDefaults, cutout:'65%', plugins:{...chartDefaults.plugins, legend:legendOpts()}}
  });

  // Chart 5 — Descartes por Tipo (cor semântica por tipo)
  destroyChart('descarte');
  const dscMap    = countBy(descartes,'tipo');
  const dscLabels = Object.keys(dscMap);
  const dscColor  = { 'Usado':C.blue, 'Novo':C.green, 'Defeito':C.red, 'Obsoleto':C.amber, 'Outros':C.slate };
  const c5 = document.getElementById('chartDescarte');
  if(c5) _corpCharts['descarte'] = new Chart(c5, {
    type:'pie',
    data:{ labels:dscLabels, datasets:[{ data:dscLabels.map(k=>dscMap[k]), backgroundColor:dscLabels.map((k,i)=>dscColor[k]||SEQ[i%SEQ.length]), borderWidth:0, hoverOffset:8 }] },
    options:{...chartDefaults, plugins:{...chartDefaults.plugins, legend:legendOpts()}}
  });

  // Chart 6 — Desembalagem (Entrada verde / Saída vermelho / Coleta âmbar)
  destroyChart('desemb');
  const c6 = document.getElementById('chartDesemb');
  if(c6) _corpCharts['desemb'] = new Chart(c6, {
    type:'bar',
    data:{ labels:['Entrada','Saída','Coleta'], datasets:[{ label:'Registros', data:[desEnt.length||2,desSai.length||3,desCol.length||1], backgroundColor:[C.green, C.red, C.amber], borderRadius:6, borderSkipped:false }] },
    options:{...chartDefaults, scales:scalesXY}
  });

  // ════════════════════════════════════════════
  // MÓDULO 3 — CONNMAP
  // ════════════════════════════════════════════
  const cmEl = document.getElementById('corpConnMapKpis');
  if(cmEl) cmEl.innerHTML = [
    kpiCard('🗺️', cmTotal,   'Total de Registros', connMap.length?'dados reais':'sem dados',                C.blue),
    kpiCard('➡️', cmA,       'Sentido A',           cmTotal?Math.round(cmA/Math.max(cmTotal,1)*100)+'%':'—', C.red),
    kpiCard('⬅️', cmB,       'Sentido B',           cmTotal?Math.round(cmB/Math.max(cmTotal,1)*100)+'%':'—', C.teal),
    kpiCard('🏭', cmVendors,  'Fabricantes',         'distintos no mapa',                                    C.purple),
    kpiCard('🗄️', cmRacks,    'Racks Mapeados',      'identificados',                                        C.orange),
    kpiCard('🔌', [...new Set(connMap.map(r=>r.iface).filter(Boolean))].length, 'Interfaces', 'tipos distintos', C.green),
  ].join('');

  setBadge('badgeConnMap', cmVendors+' fabricantes');
  setBadge('badgeConnDir', cmA+' A · '+cmB+' B');

  const connProgEl = document.getElementById('corpConnProgresso');
  if(connProgEl && cmTotal>0) connProgEl.innerHTML = `
    <div class="corp-prog-label"><span>Sentido A</span><span>${cmA} (${Math.round(cmA/cmTotal*100)}%)</span></div>
    <div class="corp-prog-track"><div class="corp-prog-fill" style="width:${Math.round(cmA/cmTotal*100)}%;background:${C.red}"></div></div>
    <div class="corp-prog-label" style="margin-top:8px"><span>Sentido B</span><span>${cmB} (${Math.round(cmB/cmTotal*100)}%)</span></div>
    <div class="corp-prog-track"><div class="corp-prog-fill" style="width:${Math.round(cmB/cmTotal*100)}%;background:${C.teal}"></div></div>`;

  // Chart 7 — ConnMap por Vendor (cada barra cor própria)
  destroyChart('connmap');
  const cmRows    = connMap.length ? connMap : [{vendor:'Cisco',direction:'A'},{vendor:'Cisco',direction:'B'},{vendor:'Huawei',direction:'A'},{vendor:'Nokia',direction:'B'},{vendor:'Cisco',direction:'A'},{vendor:'Huawei',direction:'B'}];
  const cmVMap    = cmRows.reduce((acc,r)=>{ const v=r.vendor||'Outros'; acc[v]=(acc[v]||0)+1; return acc; },{});
  const cmVLabels = Object.keys(cmVMap).sort((a,b)=>cmVMap[b]-cmVMap[a]).slice(0,8);
  const c7 = document.getElementById('chartConnMap');
  if(c7) _corpCharts['connmap'] = new Chart(c7, {
    type:'bar',
    data:{ labels:cmVLabels, datasets:[{ data:cmVLabels.map(k=>cmVMap[k]), backgroundColor:cmVLabels.map((_,i)=>SEQ[i%SEQ.length]), borderRadius:6, borderSkipped:false }] },
    options:{...chartDefaults, scales:scalesXY}
  });

  // Chart 8 — Sentido A vs B vs Indefinido
  destroyChart('conndir');
  const cmAf  = cmRows.filter(r=>r.direction==='A').length;
  const cmBf  = cmRows.filter(r=>r.direction==='B').length;
  const cmOth = cmRows.length - cmAf - cmBf;
  const c8 = document.getElementById('chartConnDir');
  if(c8) _corpCharts['conndir'] = new Chart(c8, {
    type:'doughnut',
    data:{ labels:['Sentido A','Sentido B','Indefinido'], datasets:[{ data:[cmAf,cmBf,cmOth], backgroundColor:[C.red, C.teal, C.silver], borderWidth:0, hoverOffset:8 }] },
    options:{...chartDefaults, cutout:'65%', plugins:{...chartDefaults.plugins, legend:legendOpts()}}
  });

  // Chart 9 — Interface UTP/MULT/MONO
  destroyChart('conniface');
  const cmIfaceMap = cmRows.reduce((acc,r)=>{ const v=r.iface||'Não def.'; acc[v]=(acc[v]||0)+1; return acc; },{});
  const cmIfaceLabels = Object.keys(cmIfaceMap).sort((a,b)=>cmIfaceMap[b]-cmIfaceMap[a]);
  setBadge('badgeConnIface', cmIfaceLabels.length+' tipos');
  const c9 = document.getElementById('chartConnIface');
  if(c9) _corpCharts['conniface'] = new Chart(c9, {
    type:'doughnut',
    data:{ labels:cmIfaceLabels, datasets:[{ data:cmIfaceLabels.map(k=>cmIfaceMap[k]), backgroundColor:[C.blue, C.orange, C.green, C.red, C.slate], borderWidth:0, hoverOffset:8 }] },
    options:{...chartDefaults, cutout:'65%', plugins:{...chartDefaults.plugins, legend:legendOpts()}}
  });

  // ── Sync hidden elements ──
  ['dashTotal','dashAtivos','dashGarantia','dashEol','dashEos'].forEach((id,i)=>{ const el=document.getElementById(id); if(el) el.textContent=[srvTotal,srvGar,srvFora,srvEol,srvEos][i]; });
  const dfEl=document.getElementById('dashFornecedor'); if(dfEl) dfEl.innerHTML=Object.entries(countBy(servers,'fornecedor')).map(([k,v])=>`<div><strong>${esc(k)}</strong>: ${v}</div>`).join('');
  ['robo_totalReg','robo_pendentes','robo_concluidas','robo_canceladas'].forEach((id,i)=>{ const el=document.getElementById(id); if(el) el.textContent=[roboTotal,roboPend,roboConcl,roboCanc][i]; });
  ['kpiTotal','kpiFiltradas','kpiA','kpiB'].forEach((id,i)=>{ const el=document.getElementById(id); if(el) el.textContent=[cmTotal,cmTotal,cmA,cmB][i]; });
}

// ===== END CORPORATE DASHBOARD =====

// ===== ENDEREÇOS DAS LOCALIDADES DATA CENTER =====
const DECL_ADDR_KEY = 'decl_addr_localidades';

const DECL_ADDR_DEFAULTS = [
  {site:'DATA CENTER BRASÍLIA', endereco:'', numero:'', complemento:'', bairro:'', cep:'', cidade:'', uf:'', telefone:''},
  {site:'SEDE HENRI DUNAT',     endereco:'', numero:'', complemento:'', bairro:'', cep:'', cidade:'', uf:'', telefone:''},
  {site:'DATA CENTER LAPA',     endereco:'', numero:'', complemento:'', bairro:'', cep:'', cidade:'', uf:'', telefone:''},
  {site:'DATA CENTER INGLESES', endereco:'', numero:'', complemento:'', bairro:'', cep:'', cidade:'', uf:'', telefone:''},
  {site:'DATA CENTER VITÓRIA',  endereco:'', numero:'', complemento:'', bairro:'', cep:'', cidade:'', uf:'', telefone:''},
  {site:'DATA CENTER MACKENZIE',endereco:'', numero:'', complemento:'', bairro:'', cep:'', cidade:'', uf:'', telefone:''},
];

function declAddrGet() {
  try { return JSON.parse(localStorage.getItem(DECL_ADDR_KEY)) || DECL_ADDR_DEFAULTS.map(r=>({...r})); }
  catch(e) { return DECL_ADDR_DEFAULTS.map(r=>({...r})); }
}

function declAddrRender() {
  const tbody = document.getElementById('declAddrTbody');
  if (!tbody) return;
  const rows = declAddrGet();
  const fields = ['site','endereco','numero','complemento','bairro','cep','cidade','uf','telefone'];
  tbody.innerHTML = '';
  rows.forEach((row, ri) => {
    const tr = document.createElement('tr');
    const cells = fields.map(f => {
      const td = document.createElement('td');
      const inp = document.createElement('input');
      inp.type = 'text';
      inp.value = row[f] || '';
      inp.placeholder = '';
      inp.dataset.field = f;
      inp.dataset.row = ri;
      inp.style.cssText = 'border:none;background:transparent;width:100%;padding:3px 5px;font-size:11px;min-height:unset;height:auto;color:var(--text);';
      inp.addEventListener('focus', function(){ this.style.outline='1px solid var(--primary)'; this.style.borderRadius='3px'; });
      inp.addEventListener('blur', function(){ this.style.outline='none'; });
      td.appendChild(inp);
      return td;
    });
    cells.forEach(td => tr.appendChild(td));
    // delete button
    const tdDel = document.createElement('td');
    tdDel.style.textAlign = 'center';
    const btnDel = document.createElement('button');
    btnDel.type = 'button';
    btnDel.title = 'Excluir linha';
    btnDel.textContent = '🗑️';
    btnDel.style.cssText = 'background:none;border:none;cursor:pointer;font-size:14px;padding:0;min-height:unset;color:var(--bad);';
    btnDel.onclick = () => { if(confirm('Excluir esta linha?')){ const d=declAddrGet(); d.splice(ri,1); localStorage.setItem(DECL_ADDR_KEY,JSON.stringify(d)); declAddrRender(); }};
    tdDel.appendChild(btnDel);
    tr.appendChild(tdDel);
    tbody.appendChild(tr);
  });
}

function declAddrAddRow() {
  const d = declAddrGet();
  d.push({site:'',endereco:'',numero:'',complemento:'',bairro:'',cep:'',cidade:'',uf:'',telefone:''});
  localStorage.setItem(DECL_ADDR_KEY, JSON.stringify(d));
  declAddrRender();
  // scroll to last row
  const tbody = document.getElementById('declAddrTbody');
  if(tbody){ const last = tbody.lastElementChild; if(last) last.querySelector('input').focus(); }
}

function declAddrSave() {
  const tbody = document.getElementById('declAddrTbody');
  if (!tbody) return;
  const rows = Array.from(tbody.querySelectorAll('tr')).map(tr => {
    const obj = {};
    tr.querySelectorAll('input[data-field]').forEach(inp => { obj[inp.dataset.field] = inp.value; });
    return obj;
  });
  localStorage.setItem(DECL_ADDR_KEY, JSON.stringify(rows));
  alert('✅ Endereços salvos com sucesso!');
}

// Render on modal open — hook into existing gerarDeclaracaoTransporte
(function(){
  const _orig = window.gerarDeclaracaoTransporte;
  if(typeof _orig === 'function'){
    window.gerarDeclaracaoTransporte = function(){
      _orig.apply(this, arguments);
      setTimeout(declAddrRender, 50);
    };
  }
  // Also render when modal becomes visible
  const obs = new MutationObserver(() => {
    const m = document.getElementById('modalDeclaracao');
    if(m && m.style.display !== 'none' && m.classList.contains('open')){
      declAddrRender();
    }
  });
  document.addEventListener('DOMContentLoaded', () => {
    const m = document.getElementById('modalDeclaracao');
    if(m) obs.observe(m, {attributes:true, attributeFilter:['style','class']});
    declAddrRender();
  });
})();
// ===== END ENDEREÇOS =====

