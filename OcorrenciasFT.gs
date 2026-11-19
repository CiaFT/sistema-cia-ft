/** APP FT V1.4.6 — Ocorrências, folgas do Pel, auditoria e liberação excepcional de prazo.
 * Arquivo complementar. Adicionar uma vez ao Apps Script existente.
 * Dados de indiciados/sindicados nunca compõem o ranking nem o arquivo público HTML.
 */
let FT_OC_CACHE_ = null;
let FT_OC_EVENTOS_CACHE_ = null;
let FT_OC_PEDIDO_FF_ = null;
let FT_OC_CREDITOS_INDICE_ = null;
let FT_OC_RESUMOS_FF_ = {};
let FT_OC_INICIO_CACHE_ = null;
let FT_OC_EFETIVO_CACHE_ = null;
let FT_OC_MIGRACAO_CHECADA_ = false;
function ocInvalidarFF_() {FT_OC_CREDITOS_INDICE_=null;FT_OC_RESUMOS_FF_={};}
function ocEfetivo_() {return FT_OC_EFETIVO_CACHE_||(FT_OC_EFETIVO_CACHE_=listarEfetivoAtualMestre_().sort(compararOrdemHierarquicaPm_));}
// Reconhece decisões diretas do Cmt e decisões delegadas gravadas pelo servidor.
function ocAuditoriaDecisaoFF_(o) {
  const historico=o.auditoria||[];
  for(let i=historico.length-1;i>=0;i--){
    const a=historico[i];
    if(!a)continue;
    const perfil=ocPerfil_({perfilAcesso:a.perfil||''}),acao=String(a.acao||'');
    const cap=ocCap_({perfilAcesso:perfil});
    if(!cap.validarFF)continue;
    const regra=cap.ffDelegada
      ? /^FF_(APROVADA|RECUSADA) — realizado pelo Gestor, sob autorização do Cmt Cia FT(?::|$)/
      : /^FF_(APROVADA|RECUSADA)(?::|$)/;
    const m=acao.match(regra);
    if(m)return {estado:m[1],registro:a};
  }
  return null;
}
function ocAplicarRegraFF_(o) {
  if(o.origem==='EXCEL_20261104'){
    const decisao=ocAuditoriaDecisaoFF_(o);
    o.ffDecisao=decisao?decisao.estado:'PENDENTE';
    o.ffAprovada=o.ffDecisao==='APROVADA';
  }
  return o;
}
function ocTexto_(v) { return String(v == null ? '' : v).trim(); }
function ocRe_(v) { return normalizarLoginRE_(v); }
function ocPerfil_(p) { return normalizarTexto_(p.perfilAcesso || '').replace(/[\s-]+/g,'_'); }
function ocCap_(p) {
  const r=ocPerfil_(p),enc=['ENC_ADM','ENC_ADM_FT','ENCARREGADO_ADM_FT','ENCARREGADO_ADMINISTRATIVO'].indexOf(r)>=0;
  const cmt=['CMT_CIA','CMT_CIA_FT','COMANDANTE_CIA_FT'].indexOf(r)>=0;
  const gestor=['GESTOR','GESTOR_FT','GESTOR_TEC_FT','GESTOR_TECNICO_FT'].indexOf(r)>=0;
  return {validar:enc||cmt||gestor,validarFF:cmt||gestor,reverterFF:gestor,importar:enc||gestor,consultarDetalhes:enc||cmt||gestor,ffDelegada:gestor&&!cmt};
}
function ocAuth_(params) {
  const a=autenticarUsuarioAtivoIcc_(params||{});
  if(!a.ok) throw new Error(a.erro);
  return {pm:a.policial,cap:ocCap_(a.policial)};
}
function ocISOValida_(v) {
  const t=ocTexto_(v),m=t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if(!m)return false;
  const d=new Date(Date.UTC(Number(m[1]),Number(m[2])-1,Number(m[3])));
  return d.toISOString().slice(0,10)===t;
}
function ocAba_() { return obterOuCriarAbaComCabecalho_('OCORRENCIAS_FT',['ID','CHAVE_DATA_BOPM','STATUS','DATA','BOPM','DADOS_JSON','ATUALIZADO_EM']); }
function ocConfig_() {
  if(FT_OC_INICIO_CACHE_!==null)return FT_OC_INICIO_CACHE_;
  const prop=PropertiesService.getScriptProperties();
  let inicio=prop.getProperty('FT_OCORRENCIAS_INICIO_FF');
  if(!inicio) {inicio=String(Date.now());prop.setProperty('FT_OCORRENCIAS_INICIO_FF',inicio);}
  return FT_OC_INICIO_CACHE_=Number(inicio);
}
/** Migração única da antiga concessão automática. Um único lote de escrita,
 * sem reimportação. Aprovações reais do Cmt e gozos anteriores são preservados.
 */
function ocRevisarImportadasUmaVez_() {
  if(FT_OC_MIGRACAO_CHECADA_)return;
  const props=PropertiesService.getScriptProperties(),chave='FT_OC_FF_REVISAO_CMT_V2';
  if(props.getProperty(chave)==='OK'){FT_OC_MIGRACAO_CHECADA_=true;return;}
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    if(props.getProperty(chave)==='OK'){FT_OC_MIGRACAO_CHECADA_=true;return;}
    const aba=obterAbaSeExistir_('OCORRENCIAS_FT');
    if(aba&&aba.getLastRow()>=2){
      const valores=aba.getRange(2,6,aba.getLastRow()-1,1).getValues();
      const registros=valores.map(function(r){return r[0]?JSON.parse(String(r[0])):null;});
      const alvos=registros.filter(function(o){return o&&o.origem==='EXCEL_20261104'&&!ocAuditoriaDecisaoFF_(o)&&(o.ffAprovada||o.ffDecisao==='APROVADA');});
      if(alvos.length){
        const usos = {};
        const lancamentosAtivos = listarLancamentosAtivosFolgas_('');
        for (let i = 0; i < lancamentosAtivos.length; i++) {
          const evento = lancamentosAtivos[i];
          if (evento.codigo === 'FF' && evento.ocCreditoId) {
            const chaveUso = evento.reNormalizado + '|' + evento.ocCreditoId;
            usos[chaveUso] = evento.dataISO;
          }
        }
        const agora=new Date().toISOString();
        registros.forEach(function(o,i){
          if(!o||alvos.indexOf(o)<0)return;
          o.ffDecisaoAnteriorImportacao=o.ffDecisao;o.ffAprovada=false;o.ffDecisao='PENDENTE';
          (o.equipe||[]).forEach(function(p){if(!p.ffGozadaEm&&usos[ocRe_(p.re)+'|'+o.id])p.ffGozadaEm=usos[ocRe_(p.re)+'|'+o.id];});
          o.auditoria=o.auditoria||[];o.auditoria.push({acao:'REVISAO_FF_IMPORTADA_AGUARDANDO_CMT',dataHora:agora,re:'',nome:'Ajuste da regra de FF importadas',perfil:'SISTEMA'});
          const json=JSON.stringify(o);if(json.length>45000)throw new Error('Histórico excede o limite na ocorrência '+o.bopm);
          valores[i]=[json];
        });
        aba.getRange(2,6,valores.length,1).setValues(valores);FT_OC_CACHE_=null;ocInvalidarFF_();
      }
    }
    props.setProperty(chave,'OK');FT_OC_MIGRACAO_CHECADA_=true;
  }finally{lock.releaseLock();}
}

function ocRegistros_() {
  if(FT_OC_CACHE_)return FT_OC_CACHE_;
  const aba=obterAbaSeExistir_('OCORRENCIAS_FT');
  if(!aba||aba.getLastRow()<2)return FT_OC_CACHE_=[];
  return FT_OC_CACHE_=aba.getRange(2,6,aba.getLastRow()-1,1).getValues().map(function(r,i){
    if(!r[0])return null;
    let d;try{d=JSON.parse(String(r[0]));}catch(e){throw new Error('Registro de ocorrência inválido na linha '+(i+2));}
    d._linha=i+2;return ocAplicarRegraFF_(d);
  }).filter(Boolean);
}
function ocPersistir_(d) {
  const aba=ocAba_(),linha=d._linha||aba.getLastRow()+1;
  const copia=Object.assign({},d);delete copia._linha;
  const json=JSON.stringify(copia);
  if(json.length>45000)throw new Error('Registro excede o tamanho permitido. Reduza o texto das observações.');
  aba.getRange(linha,1,1,7).setValues([[d.id,d.data+'|'+d.bopm.toUpperCase(),d.status,d.data,"'"+d.bopm,json,new Date()]]);
  FT_OC_CACHE_=null;ocInvalidarFF_();return d;
}
function ocAuditar_(d,ator,acao) {
  d.auditoria=d.auditoria||[];d.auditoria.push({acao:acao,dataHora:new Date().toISOString(),re:ator.re,nome:ator.nomeCompleto||ator.nomeGuerra,perfil:ator.perfilAcesso});
}
function ocPmSeguro_(p) { return {re:p.re,posto:p.postoGrad||'',nome:p.nomeCompleto||p.nomeGuerra||'',nomeGuerra:p.nomeGuerra||'',pelotao:grupoFrequencia_(p)}; }
function ocViaturaDisponivel_(v) {
  if(!v || !ocTexto_(v.prefixo))return false;
  const status=normalizarTexto_(v.statusApp||'');
  if(status!=='ATIVA')return false;
  const situacao=normalizarTexto_(v.situacao||'').replace(/[^A-Z0-9]+/g,' ').trim();
  const opVerao=/^(?:OP|OPERACAO) VERAO(?:$| )/.test(situacao);
  const bloqueadaResenha=typeof situacaoViaturaBloqueadaParaResenha_==='function'
    ? situacaoViaturaBloqueadaParaResenha_(situacao)
    : ['PROJETO BASICO','SINDICANCIA','DESCARGA','BAIXADA','BAIXADA PEL'].indexOf(situacao)!==-1;
  return !opVerao&&!bloqueadaResenha;
}
function ocCatalogos_(auth) {
  const viaturas=listarViaturasConsulta_();
  return {painelUnificadoFF:true,versaoModulo:'V1.4',usuario:ocPmSeguro_(auth.pm),capacidades:auth.cap,efetivo:ocEfetivo_().map(ocPmSeguro_),
    viaturas:[].concat(viaturas.tatico||[],viaturas.rocam||[],viaturas.outras||[]).filter(ocViaturaDisponivel_).map(function(v){return {prefixo:v.prefixo};})};
}
function ocSalvarEnvio_(params,auth) {
  const f=parseJsonParam_(params.campos,{}),hoje=Utilities.formatDate(new Date(),FT_API.timezone,'yyyy-MM-dd');
  if(!ocISOValida_(f.data)||f.data>hoje)throw new Error('Informe uma data válida, até hoje.');
  if(!/^([01]\d|2[0-3]):[0-5]\d$/.test(ocTexto_(f.qtr)))throw new Error('Informe o horário da ocorrência.');
  ['qth','bopm','natureza','vtr'].forEach(function(k){if(!ocTexto_(f[k]))throw new Error('Campo obrigatório: '+k.toUpperCase());});
  if(!['ADM','1º PEL','2º PEL','DEJEM'].includes(f.pel))throw new Error('Selecione o pelotão ou DEJEM.');
  if(!Array.isArray(f.equipe)||!f.equipe.length||f.equipe.length>100)throw new Error('Selecione ao menos um policial da equipe.');
  const roster=ocEfetivo_(),vistos={};
  const equipe=f.equipe.map(function(re){const base=ocRe_(re);if(vistos[base])throw new Error('O mesmo policial foi selecionado mais de uma vez.');vistos[base]=true;
    const p=roster.filter(function(pm){return ocRe_(pm.re)===base;})[0];if(!p)throw new Error('Integrante não localizado no efetivo: '+re);return ocPmSeguro_(p);});
  const frota=listarViaturasConsulta_(),vtrs=[].concat(frota.tatico||[],frota.rocam||[],frota.outras||[]);
  if(!vtrs.some(function(v){return String(v.prefixo)===String(f.vtr)&&ocViaturaDisponivel_(v);}))throw new Error('Viatura indisponível para inclusão de ocorrência. Selecione uma viatura ativa, fora de OP VERÃO e não bloqueada para a Resenha.');
  const pessoas=function(itens){if(!Array.isArray(itens)||itens.length>200)throw new Error('Relação de pessoas inválida.');return itens.filter(function(p){return ocTexto_(p.nome)||ocTexto_(p.rg);}).map(function(p){if(!ocTexto_(p.nome))throw new Error('Informe o nome na relação de pessoas.');return {nome:ocTexto_(p.nome),rg:ocTexto_(p.rg)};});};
  const request=ocTexto_(params.requestId);if(!/^[A-Za-z0-9-]{16,80}$/.test(request))throw new Error('Protocolo do envio inválido. Atualize a página.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const todos=ocRegistros_(),repetida=todos.filter(function(o){return o.requestId===request;})[0];
    if(repetida){if(ocRe_(repetida.responsavel.re)!==ocRe_(auth.pm.re))throw new Error('Protocolo já utilizado.');return {id:repetida.id,idempotente:true};}
    const chave=ocTexto_(f.data)+'|'+ocTexto_(f.bopm).toUpperCase();
    if(todos.some(function(o){return o.data+'|'+o.bopm.toUpperCase()===chave;}))throw new Error('Já existe ocorrência com esse BOPM nesta data. Consulte o registro antes de reenviar.');
    const d={id:Utilities.getUuid(),requestId:request,responsavel:ocPmSeguro_(auth.pm),enviadaEm:new Date().toISOString(),pel:f.pel,data:f.data,qtr:ocTexto_(f.qtr),qth:ocTexto_(f.qth),bopm:ocTexto_(f.bopm),bopc:ocTexto_(f.bopc),dp:ocTexto_(f.dp),natureza:ocTexto_(f.natureza),vtr:ocTexto_(f.vtr),equipe:equipe,apoio:(f.apoio||[]).map(ocTexto_).filter(Boolean),indiciados:pessoas(f.indiciados||[]),sindicados:pessoas(f.sindicados||[]),apreensao:ocTexto_(f.apreensao),status:'PENDENTE',ffDecisao:'PENDENTE',ffAprovada:false,origem:'PORTAL'};
    ocAuditar_(d,auth.pm,'ENVIAR');ocPersistir_(d);return {id:d.id,mensagem:'Ocorrência enviada para validação.'};
  }finally{lock.releaseLock();}
}
function ocValidar_(params,auth) {
  if(!['FF','OCORRENCIA'].includes(params.tipo)||!['APROVAR','RECUSAR'].includes(params.decisao))throw new Error('Decisão inválida.');
  const ff=params.tipo==='FF',aprovar=params.decisao==='APROVAR';
  if(ff?!auth.cap.validarFF:!auth.cap.validar)throw new Error(ff?'A decisão de FF é restrita ao Cmt Cia ou ao Gestor sob autorização.':'Somente Enc Adm, Cmt Cia ou Gestor podem validar a ocorrência.');
  const delegada=ff&&auth.cap.ffDelegada===true;
  if(delegada&&params.autorizacaoCmt!=='SIM')throw new Error('Confirme a autorização do Cmt Cia FT.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const d=ocRegistros_().filter(function(o){return o.id===params.id;})[0];if(!d)throw new Error('Ocorrência não localizada.');
    const atual=ff?d.ffDecisao:d.status,final=ff?(aprovar?'APROVADA':'RECUSADA'):(aprovar?'VALIDADA':'RECUSADA');
    if(atual===final)return {id:d.id,idempotente:true};
    if(atual!=='PENDENTE')throw new Error('Esta decisão já foi registrada.');
    if(ff&&d.status!=='VALIDADA')throw new Error('Valide a ocorrência antes de decidir a FF.');
    const motivo=ocTexto_(params.motivo);if(!aprovar&&motivo.length<5)throw new Error('Informe o motivo da recusa.');
    if(ff){d.ffDecisao=final;d.ffAprovada=aprovar;}else d.status=final;
    ocAuditar_(d,auth.pm,(ff?'FF_':'OCORRENCIA_')+final+(delegada?' — realizado pelo Gestor, sob autorização do Cmt Cia FT':'')+(motivo?': '+motivo:''));ocPersistir_(d);return {id:d.id,mensagem:'Decisão registrada.'};
  }finally{lock.releaseLock();}
}
/** Alocação pura: cada FF do controle usa um crédito. Importação com data de
 * gozo é preservada e não é descontada duas vezes se existir evento legado.
 */
function ocReverterRecusaFF_(params,auth) {
  if(!auth.cap.reverterFF)throw new Error('Somente o Gestor pode reverter a recusa de uma FF.');
  if(params.autorizacaoCmt!=='SIM')throw new Error('Confirme a autorização do Cmt Cia FT para conceder a FF.');
  const motivo=ocTexto_(params.motivo);
  if(motivo.length<5)throw new Error('Informe a justificativa da reversão (mínimo de 5 caracteres).');
  const id=ocTexto_(params.id);
  if(!id)throw new Error('Ocorrência não informada.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const d=ocRegistros_().filter(function(o){return o.id===id;})[0];
    if(!d)throw new Error('Ocorrência não localizada.');
    if(d.status!=='VALIDADA')throw new Error('Só é possível conceder a FF de uma ocorrência já validada.');
    if(d.ffDecisao!=='RECUSADA')throw new Error(d.ffDecisao==='APROVADA'?'Esta FF já está concedida.':'A FF não está recusada; atualize a página e confira a situação.');
    d.ffDecisao='APROVADA';
    d.ffAprovada=true;
    ocAuditar_(d,auth.pm,'FF_APROVADA — realizado pelo Gestor, sob autorização do Cmt Cia FT: reversão da recusa anterior. Justificativa: '+motivo);
    ocPersistir_(d);
    return {id:d.id,mensagem:'Recusa revertida. A FF foi concedida para os integrantes da equipe e os créditos foram recalculados.'};
  }finally{lock.releaseLock();}
}
function ocAlocarFF_(creditos,eventos,dataFinal) {
  const copia=creditos.map(function(c){return Object.assign({},c);}).sort(function(a,b){return a.data.localeCompare(b.data)||a.id.localeCompare(b.id);});
  const usados={},porDia={},erros=[];
  copia.forEach(function(c){if(c.gozadaEm)usados[c.id]=c.gozadaEm;});
  const reservados={};(eventos||[]).forEach(function(e){if(e.codigo==='FF'&&e.ocCreditoId)reservados[e.ocCreditoId]=e.dataISO;});
  (eventos||[]).filter(function(e){return e.codigo==='FF'&&(!dataFinal||e.dataISO<=dataFinal)&&(e.ocGerenciado||e.marcador);}).slice().sort(function(a,b){return a.dataISO.localeCompare(b.dataISO)||Number(a.linha||0)-Number(b.linha||0);}).forEach(function(e){
    const pronto=copia.filter(function(c){return c.gozadaEm===e.dataISO;})[0];
    if(pronto){porDia[e.dataISO]=pronto;return;}
    let livres=copia.filter(function(c){return !usados[c.id]&&c.data<=e.dataISO&&(!reservados[c.id]||reservados[c.id]===e.dataISO);});
    let c=e.ocCreditoId?livres.filter(function(c){return c.id===e.ocCreditoId;})[0]:livres[0];
    if(!c){erros.push({data:e.dataISO,proposta:!!e.marcador});return;}
    usados[c.id]=e.dataISO;c.gozadaEm=e.dataISO;porDia[e.dataISO]=c;
  });
  return {creditos:copia,pendentes:copia.filter(function(c){return !usados[c.id];}),porDia:porDia,usados:usados,erros:erros};
}
function ocCreditosFF_(re) {
  if(FT_OC_CREDITOS_INDICE_===null) {
    const porRe={},usos={};
    const registros=ocRegistros_(),precisaPreservarUso=registros.some(function(o){return o.origem==='EXCEL_20261104'&&!o.ffAprovada;});
    if(precisaPreservarUso)listarLancamentosAtivosFolgas_('').forEach(function(e){
      if(e.codigo==='FF'&&e.ocCreditoId)usos[e.reNormalizado+'|'+e.ocCreditoId]=e.dataISO;
    });
    registros.forEach(function(o){(o.equipe||[]).forEach(function(p){
      const chave=ocRe_(p.re),gozo=p.ffGozadaEm||usos[chave+'|'+o.id]||'';
      if(!o.ffAprovada&&!gozo)return;
      if(!porRe[chave])porRe[chave]=[];
      porRe[chave].push({id:o.id,bopm:o.bopm,data:o.data,natureza:o.natureza,gozadaEm:gozo});
    });});
    Object.keys(porRe).forEach(function(r){porRe[r].sort(function(a,b){return a.data.localeCompare(b.data)||a.id.localeCompare(b.id);});});
    FT_OC_CREDITOS_INDICE_=porRe;
  }
  return FT_OC_CREDITOS_INDICE_[ocRe_(re)]||[];
}
function ocResumoFF_(re,ate,eventos) {
  if(eventos)return ocAlocarFF_(ocCreditosFF_(re),eventos,ate);
  const chave=ocRe_(re)+'|'+(ate||'');
  if(!FT_OC_RESUMOS_FF_[chave])FT_OC_RESUMOS_FF_[chave]=ocAlocarFF_(ocCreditosFF_(re),listarLancamentosAtivosFolgas_(re),ate);
  return FT_OC_RESUMOS_FF_[chave];
}
function ocResumoOcorrencia_(o) {
  // A lista precisa apenas do resumo. Equipe, RG e auditoria ficam no detalhe.
  return {id:o.id,data:o.data,natureza:o.natureza,bopm:o.bopm,status:o.status,ffDecisao:o.ffDecisao,origem:o.origem,ffAprovada:o.ffAprovada};
}
function ocFiltrosRanking_(params) {
  if(typeof params!=='object')params={ano:params};
  params=params||{};
  const ano=ocTexto_(params.ano),periodo=ocTexto_(params.periodo)||'ANUAL';
  if(!['ANUAL','SEMESTRAL','MENSAL'].includes(periodo))throw new Error('Período inválido.');
  if(ano&&!/^[0-9]{4}$/.test(ano))throw new Error('Ano inválido.');
  if(periodo!=='ANUAL'&&!ano)throw new Error('Selecione o ano.');
  const mes=Number(params.mes),semestre=Number(params.semestre),pel=ocTexto_(params.pel);
  if(periodo==='MENSAL'&&(!Number.isInteger(mes)||mes<1||mes>12))throw new Error('Mês inválido.');
  if(periodo==='SEMESTRAL'&&![1,2].includes(semestre))throw new Error('Semestre inválido.');
  if(pel&&!['1º PEL','2º PEL','ADM'].includes(pel))throw new Error('Pelotão inválido.');
  return {ano:ano,periodo:periodo,mes:mes,semestre:semestre,pel:pel};
}
function ocNoPeriodoRanking_(o,f) {
  const mes=Number(o.data.slice(5,7));
  return (!f.ano||o.data.slice(0,4)===f.ano)&&(f.periodo!=='MENSAL'||mes===f.mes)&&(f.periodo!=='SEMESTRAL'||Math.ceil(mes/6)===f.semestre);
}
function ocFFEmAnalise_(o,p) {return o.status==='VALIDADA'&&o.ffDecisao==='PENDENTE'&&!p.ffGozadaEm;}
function ocRanking_(params) {
  const f=ocFiltrosRanking_(params),contagens={},pendentes={};
  // Uma passagem nos registros já carregados; nenhum cálculo do Controle de Folgas.
  ocRegistros_().forEach(function(o){
    if(o.status!=='VALIDADA'||!ocNoPeriodoRanking_(o,f))return;
    const vistos={};(o.equipe||[]).forEach(function(p){const r=ocRe_(p.re);if(vistos[r])return;vistos[r]=true;
      contagens[r]=(contagens[r]||0)+1;
      if(ocFFEmAnalise_(o,p))pendentes[r]=(pendentes[r]||0)+1;
    });
  });
  return ocEfetivo_().slice().sort(compararOrdemHierarquicaPm_).map(function(p,i){return Object.assign(ocPmSeguro_(p),{total:contagens[ocRe_(p.re)]||0,ffEmAnalise:pendentes[ocRe_(p.re)]||0,ordem:i});}).filter(function(p){return !f.pel||p.pelotao===f.pel;}).sort(function(a,b){return b.total-a.total||a.ordem-b.ordem;});
}
function ocPendenciasRanking_(params,auth) {
  const f=ocFiltrosRanking_(params),re=ocRe_(params.re);
  if(!re)throw new Error('Selecione o policial.');
  const dados=ocRegistros_().filter(function(o){return ocNoPeriodoRanking_(o,f)&&(o.equipe||[]).some(function(p){return ocRe_(p.re)===re&&ocFFEmAnalise_(o,p);});}).map(ocResumoOcorrencia_).sort(function(a,b){return a.data.localeCompare(b.data);});
  return {ocorrencias:dados,capacidades:auth.cap};
}
/** Resumos do painel único. Dados completos e RG somente no detalhe autorizado. */
function ocPainelFF_(params,auth) {
  const cia=params.modo==='CIA';
  if(cia&&!auth.cap.validarFF)throw new Error('As pendências da Cia são restritas ao Cmt Cia e ao Gestor.');
  const re=ocRe_(params.re||auth.pm.re);
  const periodo=params.todoHistorico==='true'?null:ocFiltrosRanking_(params);
  const creditos={};if(!cia)ocResumoFF_(re).creditos.forEach(function(c){creditos[c.id]=c;});
  const lista=ocRegistros_().filter(function(o){return (!periodo||ocNoPeriodoRanking_(o,periodo))&&(cia?(o.status==='VALIDADA'&&(o.ffDecisao==='PENDENTE'||(auth.cap.reverterFF&&o.ffDecisao==='RECUSADA'))&&(o.equipe||[]).some(function(p){return !p.ffGozadaEm;})):(o.equipe||[]).some(function(p){return ocRe_(p.re)===re;}));}).map(function(o){
    const c=creditos[o.id],gozo=c&&c.gozadaEm||'';
    const situacao=gozo?'FRUIDA':o.status==='RECUSADA'||o.ffDecisao==='RECUSADA'?'RECUSADA':o.status!=='VALIDADA'?'VALIDACAO':o.ffAprovada?'DISPONIVEL':'AGUARDANDO';
    const decisao=ocAuditoriaDecisaoFF_(o),audit=decisao&&decisao.registro;
    return Object.assign(ocResumoOcorrencia_(o),{situacao:situacao,fruidaEm:gozo,decisao:audit?{acao:audit.acao,nome:audit.nome,dataHora:audit.dataHora}:null,
      equipe:cia?(o.equipe||[]).map(function(p){return {re:p.re,posto:p.posto,nome:p.nomeGuerra||p.nome,fruidaEm:p.ffGozadaEm||''};}):undefined});
  }).sort(function(a,b){return a.data.localeCompare(b.data)||a.id.localeCompare(b.id);});
  return {ocorrencias:lista,capacidades:auth.cap,re:re,modo:cia?'CIA':'POLICIAL'};
}
function ocImportar_(params,auth) {
  if(!auth.cap.importar)throw new Error('Somente Gestor ou Enc Adm pode importar o histórico.');
  const lote=parseJsonParam_(params.lote,[]);if(!Array.isArray(lote)||lote.length>8)throw new Error('Lote inválido.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const existentes=ocRegistros_().slice();let novos=0,ignorados=0;
    lote.forEach(function(f){
      if(!/^LEG-[a-f0-9]{24}$/.test(f.id)||!ocISOValida_(f.data)||!ocTexto_(f.bopm)||!Array.isArray(f.equipe)||!f.equipe.length)throw new Error('Registro histórico incompleto.');
      if(existentes.some(function(o){return o.id===f.id;})){ignorados++;return;}
      if(existentes.some(function(o){return o.data===f.data&&o.bopm.toUpperCase()===ocTexto_(f.bopm).toUpperCase();}))throw new Error('BOPM já cadastrado nesta data: '+f.bopm);
      if(existentes.some(function(o){return (o.fontes||[]).some(function(a){return (f.fontes||[]).some(function(b){return a.aba===b.aba&&Number(a.linha)===Number(b.linha);});});}))throw new Error('Uma linha desta ocorrência já foi importada com outra identificação. Confira o BOPM '+f.bopm+' antes de continuar.');
      const membros={};f.equipe.forEach(function(p){const chave=ocRe_(p.re);if(membros[chave])throw new Error('Policial duplicado na equipe histórica.');membros[chave]=true;});
      f.equipe.forEach(function(p){if(!/^\d{6}-[0-9A-Z]$/.test(p.re))throw new Error('RE histórico inválido.');if(p.ffGozadaEm&&(!ocISOValida_(p.ffGozadaEm)||p.ffGozadaEm<f.data))throw new Error('Data de gozo histórica inválida.');});
      const d={id:f.id,data:f.data,bopm:ocTexto_(f.bopm),natureza:ocTexto_(f.natureza),qth:ocTexto_(f.qth),equipe:f.equipe.map(function(p){return {re:p.re,posto:ocTexto_(p.posto),nome:ocTexto_(p.nome),ffGozadaEm:p.ffGozadaEm||''};}),status:'VALIDADA',ffDecisao:'PENDENTE',ffAprovada:false,origem:'EXCEL_20261104',fontes:f.fontes||[],responsavel:ocPmSeguro_(auth.pm),enviadaEm:new Date().toISOString()};
      ocAuditar_(d,auth.pm,'IMPORTACAO_EXCEL_CONFIRMADA');ocPersistir_(d);existentes.push(d);novos++;
    });return {novos:novos,ignorados:ignorados};
  }finally{lock.releaseLock();}
}
function ocValIndicacaoPel_(v){const t=normalizarTexto_(v||'').replace(/[^A-Z0-9]+/g,' ').trim();if(/(^|\s)1(\s|$)|PRIMEIRO/.test(t))return '1º PEL';if(/(^|\s)2(\s|$)|SEGUNDO/.test(t))return '2º PEL';return '';}
function ocValIndicacaoEscopo_(auth,pelInformado) {
  const perfil=ocPerfil_(auth.pm),grupo=ocValIndicacaoPel_(grupoFrequencia_(auth.pm));
  const gestor=['GESTOR','GESTOR_FT','GESTOR_TEC_FT','GESTOR_TECNICO_FT'].indexOf(perfil)>=0;
  const enc=['ENC_ADM','ENC_ADM_FT','ENCARREGADO_ADM_FT','ENCARREGADO_ADMINISTRATIVO'].indexOf(perfil)>=0;
  const proprio=grupo==='1º PEL'||grupo==='2º PEL'?grupo:'';
  let pel='';
  if(gestor||enc){pel=ocValIndicacaoPel_(pelInformado);if(!pel)throw new Error('Selecione o pelotão da indicação.');}
  else if(/SGT|SARGENTO/.test(ocTexto_(auth.pm.postoGrad).toUpperCase())&&proprio){
    if(pelInformado&&ocValIndicacaoPel_(pelInformado)!==proprio)throw new Error('O Sgt só pode indicar ocorrências do próprio pelotão.');
    pel=proprio;
  }else throw new Error('Acesso restrito aos Sgts dos pelotões, Enc Adm e Gestor.');
  return {pel:pel,podeEscolherPel:gestor||enc};
}
function ocValIndicacaoSemana_(){
  const hoje=Utilities.formatDate(new Date(),FT_API.timezone,'yyyy-MM-dd'),d=new Date(hoje+'T12:00:00Z');
  const deslocamento=(d.getUTCDay()-2+7)%7;
  d.setUTCDate(d.getUTCDate()-deslocamento);const inicio=d.toISOString().slice(0,10);
  d.setUTCDate(d.getUTCDate()+6);return {inicio:inicio,fim:d.toISOString().slice(0,10)};
}
function ocValIndicacaoAba_(){return obterOuCriarAbaComCabecalho_('INDICACOES_LMP_ELOGIOS',['ID','OCORRENCIA_ID','PELOTAO','STATUS','DADOS_JSON','CRIADO_POR_RE','CRIADO_EM','ATUALIZADO_EM']);}
function ocValIndicacaoOcorrencia_(id,pel,janela){
  const o=ocRegistros_().filter(function(x){return x.id===id;})[0];
  if(!o||ocValIndicacaoPel_(o.pel)!==pel)throw new Error('Ocorrência não encontrada para o pelotão selecionado.');
  if(o.data<janela.inicio||o.data>janela.fim)throw new Error('A indicação só pode ser feita para ocorrência da semana atual ('+janela.inicio+' a '+janela.fim+').');
  if(o.status==='RECUSADA')throw new Error('Ocorrência recusada não pode receber indicação.');
  const equipe=(o.equipe||[]).filter(function(p){return ocValIndicacaoPel_(p.pelotao)===pel;}).map(function(p){return {re:ocRe_(p.re),posto:ocTexto_(p.posto),nome:ocTexto_(p.nome),nomeGuerra:ocTexto_(p.nomeGuerra)};});
  return {ocorrencia:o,equipe:equipe};
}
function ocValIndicacoes_(params,auth,salvar){
  const escopo=ocValIndicacaoEscopo_(auth,params.pel),janela=ocValIndicacaoSemana_();
  const aba=ocValIndicacaoAba_();
  if(!salvar){
    const ocorrencias=ocRegistros_().filter(function(o){return ocValIndicacaoPel_(o.pel)===escopo.pel&&o.data>=janela.inicio&&o.data<=janela.fim&&o.status!=='RECUSADA';}).map(function(o){
      const equipe=(o.equipe||[]).filter(function(p){return ocValIndicacaoPel_(p.pelotao)===escopo.pel;}).map(function(p){return {re:ocRe_(p.re),posto:ocTexto_(p.posto),nome:ocTexto_(p.nome),nomeGuerra:ocTexto_(p.nomeGuerra)};});
      return {id:o.id,data:o.data,bopm:o.bopm,natureza:o.natureza,status:o.status,equipe:equipe};
    }).filter(function(o){return o.equipe.length>0;}).sort(function(a,b){return b.data.localeCompare(a.data)||String(a.bopm).localeCompare(String(b.bopm));});
    const registros=aba.getLastRow()<2?[]:aba.getRange(2,1,aba.getLastRow()-1,8).getValues().map(function(r){
      if(r[2]!==escopo.pel)return null;let d={};try{d=JSON.parse(String(r[4]||'{}'));}catch(e){}
      return {id:r[0],ocorrenciaId:r[1],pel:r[2],status:r[3],data:d.data||'',bopm:d.bopm||'',natureza:d.natureza||'',policiais:Array.isArray(d.policiais)?d.policiais:[],criadoPorRe:r[5],criadoEm:r[6]};
    }).filter(Boolean).sort(function(a,b){return String(b.data).localeCompare(String(a.data));});
    return {pel:escopo.pel,podeEscolherPel:escopo.podeEscolherPel,semana:janela,ocorrencias:ocorrencias,registros:registros};
  }
  const id=ocTexto_(params.ocorrenciaId),selecionados=parseJsonParam_(params.policiais,[]);
  if(!Array.isArray(selecionados)||!selecionados.length)throw new Error('Selecione ao menos um policial e marque LMP e/ou Elogio individual.');
  const alvo=ocValIndicacaoOcorrencia_(id,escopo.pel,janela),mapa={};alvo.equipe.forEach(function(p){mapa[p.re]=p;});
  const vistos={};const policiais=selecionados.map(function(x){const re=ocRe_(x.re),pm=mapa[re];if(!pm||vistos[re])throw new Error('A relação de policiais não corresponde à equipe desta ocorrência.');vistos[re]=true;
    const lmp=x.lmp===true,elogio=x.elogio===true;if(!lmp&&!elogio)throw new Error('Cada policial selecionado deve ter LMP, Elogio individual ou ambos.');
    return {re:pm.re,posto:pm.posto,nome:pm.nome,nomeGuerra:pm.nomeGuerra,lmp:lmp,elogio:elogio};});
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const last=aba.getLastRow(),rows=last<2?[]:aba.getRange(2,1,last-1,8).getValues();
    const existente=rows.findIndex(function(r){return r[1]===id&&r[2]===escopo.pel;});
    if(existente>=0&&ocRe_(rows[existente][5])!==ocRe_(auth.pm.re))throw new Error('Já existe uma indicação para esta ocorrência e pelotão, registrada por outro responsável. Consulte o Enc Adm ou o Gestor para evitar duplicidade.');
    const agora=new Date().toISOString(),registro=JSON.stringify({data:alvo.ocorrencia.data,bopm:alvo.ocorrencia.bopm,natureza:alvo.ocorrencia.natureza,policiais:policiais});
    const registroId=existente>=0?rows[existente][0]:Utilities.getUuid(),linha=existente>=0?existente+2:aba.getLastRow()+1;
    aba.getRange(linha,1,1,8).setValues([[registroId,id,escopo.pel,'AGUARDANDO_FORMULARIO',registro,ocRe_(auth.pm.re),existente>=0?rows[existente][6]:agora,agora]]);
    return {id:registroId,status:'AGUARDANDO_FORMULARIO',policiais:policiais.length};
  }finally{lock.releaseLock();}
}
function ocExecutar_(acao,params) {
  try{
    const auth=ocAuth_(params);
    if(acao.indexOf('fc_adm_')===0){
      const operacao=acao.slice('fc_adm_'.length);
      if(!['dados','pedir','decidir','configurar'].includes(operacao))throw new Error('Ação de FC da ADM desconhecida.');
      return respostaOk_(epExecutar_('espaco_pel_'+operacao,Object.assign({},params,{pel:'ADM'}),auth,true));
    }
    if(acao.indexOf('folga_comp_')===0)return respostaOk_(fcExecutar_(acao,params,auth));
    const calendario=(acao==='espaco_pel_dados'&&params.semSaldo==='1')||acao==='espaco_pel_resenha'||acao.indexOf('valor_lmp_indicacoes_')===0||['espaco_pel_mensagens','espaco_pel_mensagens_contador','espaco_pel_destinatarios'].includes(acao);
    if(!calendario){ocConfig_();ocRevisarImportadasUmaVez_();}
    if(acao.indexOf('espaco_pel_')===0)return respostaOk_(epExecutar_(acao,params,auth));
    if(acao==='ocorrencias_catalogos')return respostaOk_(ocCatalogos_(auth));
    if(acao==='valor_lmp_indicacoes_listar')return respostaOk_(ocValIndicacoes_(params,auth,false));
    if(acao==='valor_lmp_indicacoes_salvar')return respostaOk_(ocValIndicacoes_(params,auth,true));
    if(acao==='ocorrencias_painel_ranking')return respostaOk_({catalogos:ocCatalogos_(auth),ranking:ocRanking_(params)});
    if(acao==='ocorrencias_enviar')return respostaOk_(ocSalvarEnvio_(params,auth));
    if(acao==='ocorrencias_validar')return respostaOk_(ocValidar_(params,auth));
    if(acao==='ocorrencias_reverter_recusa_ff')return respostaOk_(ocReverterRecusaFF_(params,auth));
    if(acao==='ocorrencias_importar')return respostaOk_(ocImportar_(params,auth));
    if(acao==='ocorrencias_resumo_importacao'){
      if(!auth.cap.importar)throw new Error('Sem permissão para conferir importação.');
      const registros=ocRegistros_().filter(function(o){return o.origem==='EXCEL_20261104';});
      return respostaOk_({ocorrencias:registros.length,participacoes:registros.reduce(function(n,o){return n+o.equipe.length;},0),gozadas:registros.reduce(function(n,o){return n+o.equipe.filter(function(p){return !!p.ffGozadaEm;}).length;},0)});
    }
    if(acao==='ocorrencias_ranking')return respostaOk_({ranking:ocRanking_(params)});
    if(acao==='ocorrencias_ff_em_analise')return respostaOk_(ocPendenciasRanking_(params,auth));
    if(acao==='ocorrencias_painel_ff')return respostaOk_(ocPainelFF_(params,auth));
    if(acao==='ocorrencias_listar'){
      const re=ocRe_(params.re||''),dados=ocRegistros_().filter(function(o){return !re||(o.equipe||[]).some(function(p){return ocRe_(p.re)===re;});}).map(function(o){return Object.assign(ocResumoOcorrencia_(o),{podeReverterFF:auth.cap.reverterFF&&o.status==='VALIDADA'&&o.ffDecisao==='RECUSADA'});}).sort(function(a,b){return b.data.localeCompare(a.data);});
      const credito=re?ocResumoFF_(re).creditos:[];
      return respostaOk_({ocorrencias:dados,creditosFF:credito,capacidades:auth.cap});
    }
    if(acao==='ocorrencias_detalhe'){
      const d=ocRegistros_().filter(function(o){return o.id===params.id;})[0];if(!d)throw new Error('Ocorrência não localizada.');
      if(!auth.cap.consultarDetalhes&&ocRe_((d.responsavel||{}).re)!==ocRe_(auth.pm.re)&&!(d.equipe||[]).some(function(p){return ocRe_(p.re)===ocRe_(auth.pm.re);}))throw new Error('Acesso restrito aos envolvidos e à administração.');
      const copia=Object.assign({},d);delete copia._linha;return respostaOk_(copia);
    }
    if(acao==='ocorrencias_prever_ff'){
      const cap=capacidadesFrequencia_(auth.pm.perfilAcesso,{});if(!cap.podeEditar)throw new Error('Sem permissão para lançar folgas.');
      const alvo=listarEfetivoControleFrequencia_().filter(function(p){return ocRe_(p.re)===ocRe_(params.re);})[0];
      if(!alvo||grupoFrequencia_(alvo)==='ADM')throw new Error('Lançamento de FF bloqueado para ADM.');
      if(!ocISOValida_(params.data))throw new Error('Data inválida.');
      const eventos=listarLancamentosAtivosFolgas_(params.re).filter(function(e){return e.dataISO!==params.data;});
      eventos.push({codigo:'FF',dataISO:params.data,marcador:'PREVIA',linha:999999999});
      const estado=ocResumoFF_(params.re,'',eventos),credito=estado.porDia[params.data];
      if(!credito)throw new Error('Não há FF concedida disponível para esta data.');
      return respostaOk_({selecionada:credito,pendentes:estado.creditos.filter(function(c){return !c.gozadaEm||c.id===credito.id;})});
    }
    throw new Error('Ação de ocorrências não reconhecida.');
  }catch(e){return respostaFalha_(e.message||String(e));}
}
const FT_OC_EXECUTAR_ANTERIOR_=executarAcao_;
executarAcao_=function(acao,params){return /^(ocorrencias_|valor_lmp_indicacoes_|espaco_pel_|folga_comp_|fc_adm_)/.test(String(acao).toLowerCase())?ocExecutar_(String(acao).toLowerCase(),params||{}):FT_OC_EXECUTAR_ANTERIOR_(acao,params);};

// Integração FF: mantém o cálculo de FM existente e associa os créditos dentro
// do mesmo bloqueio e da mesma reversão já usados pela gravação diária.
listarLancamentosAtivosFolgas_=function(re){
  if(!FT_OC_EVENTOS_CACHE_){
    const aba=obterAbaObrigatoria_(FT_API.abas.escalasFolgas),n=aba.getLastRow();
    const linhas=n>=2?aba.getRange(2,1,n-1,25).getValues():[],inicio=ocConfig_(),mapa={};
    linhas.forEach(function(row,i){
      const chaveRe=ocRe_(row[3]),data=converterDataControleFolgasParaIso_(row[2]);
      if(!chaveRe||!data||normalizarTexto_(row[23]||'')!=='ATIVO')return;
      const m=String(row[24]||'').match(/FF_OCORRENCIA:([A-Za-z0-9-]+)/);
      mapa[chaveRe+'|'+data]={linha:i+2,id:row[0]||'',reNormalizado:chaveRe,dataISO:data,codigo:String(row[9]||row[10]||'').trim().toUpperCase(),ocCreditoId:m?m[1]:'',ocGerenciado:!!m||(row[1] instanceof Date&&row[1].getTime()>=inicio)};
    });
    FT_OC_EVENTOS_CACHE_=Object.keys(mapa).map(function(k){return mapa[k];}).sort(ordenarEventosControleFolgas_);
  }
  const alvo=re?ocRe_(re):'';
  return FT_OC_EVENTOS_CACHE_.filter(function(e){return !alvo||e.reNormalizado===alvo;}).map(function(e){return Object.assign({},e);});
};
const FT_OC_RECONSTRUIR_ANTERIOR_=reconstruirEstadoFolgas_;
reconstruirEstadoFolgas_=function(re,iniciais,eventos,ate){
  const r=FT_OC_RECONSTRUIR_ANTERIOR_(re,iniciais,eventos,ate);if(!r.ok)return r;
  // A leitura mensal e um EO/FM comum não precisam recalcular FF nesta etapa.
  const alteracao=(eventos||[]).find(e=>e.marcador&&e.codigo&&e.codigo!=='FOL');
  if(alteracao&&fcBloqueioOficial_(ocRe_(re),alteracao.dataISO))return {ok:false,erro:'Há uma FC aprovada ou gozada para essa data. Confira em Folgas Pendentes.',anotacoes:[]};
  const proposta=(eventos||[]).filter(function(e){return e.marcador&&e.codigo==='FF';})[0];
  if(!proposta)return r;
  const ff=ocResumoFF_(re,ate,eventos);
  if(ff.erros.some(function(e){return e.proposta;}))return {ok:false,erro:'Não há FF aprovada pelo Cmt Cia disponível nesta data.',anotacoes:[]};
  if(FT_OC_PEDIDO_FF_&&ocRe_(re)===FT_OC_PEDIDO_FF_.re){const c=ff.porDia[FT_OC_PEDIDO_FF_.data];if(!c||c.id!==FT_OC_PEDIDO_FF_.id)return {ok:false,erro:'A FF disponível mudou. Recarregue e confirme novamente.',anotacoes:[]};}
  return r;
};
const FT_OC_RECALCULAR_ANTERIOR_=recalcularPersistirFolgas_;
recalcularPersistirFolgas_=function(re){
  FT_OC_EVENTOS_CACHE_=null;ocInvalidarFF_();
  const r=FT_OC_RECALCULAR_ANTERIOR_(re),eventos=listarLancamentosAtivosFolgas_(re);
  const vincular=eventos.filter(function(e){return e.codigo==='FF'&&e.ocGerenciado&&!e.ocCreditoId;});
  if(!vincular.length)return r;
  const ff=ocResumoFF_(re,'',eventos),aba=obterAbaObrigatoria_(FT_API.abas.escalasFolgas);
  vincular.forEach(function(e){const c=ff.porDia[e.dataISO];if(c){
    aba.getRange(e.linha,25).setValue('FF_OCORRENCIA:'+c.id+' | BOPM '+c.bopm+' | '+c.data);
    FT_OC_EVENTOS_CACHE_.forEach(function(atual){if(atual.linha===e.linha)atual.ocCreditoId=c.id;});
  }});
  ocInvalidarFF_();return r;
};
const FT_OC_SALVAR_FOLGAS_ANTERIOR_=apiSalvarControleFolgas_;
apiSalvarControleFolgas_=function(params){
  const ff=String(params.valor||'').toUpperCase()==='FF';
  if(ff&&!params.ffOcorrenciaId)return respostaFalha_('Confirme a FF mais antiga no aviso antes de salvar.');
  FT_OC_PEDIDO_FF_=ff?{re:ocRe_(params.re),data:params.competencia+'-'+String(params.dia).padStart(2,'0'),id:String(params.ffOcorrenciaId)}:null;
  try{return FT_OC_SALVAR_FOLGAS_ANTERIOR_(params);}finally{FT_OC_PEDIDO_FF_=null;}
};
const FT_OC_DADOS_FOLGAS_ANTERIOR_=apiDadosControleFolgas_;
apiDadosControleFolgas_=function(params){const r=FT_OC_DADOS_FOLGAS_ANTERIOR_(params);if(r&&r.ok&&r.dados){Object.keys(r.dados.saldos||{}).forEach(function(re){const ff=ocResumoFF_(re);r.dados.saldos[re].creditosFF=ff.pendentes;r.dados.saldos[re].folgasFlagrante=ff.pendentes.length;});}return r;};
const FT_OC_PENDENTES_ANTERIOR_=listarFolgasPendentesEscalasFolgas_;
listarFolgasPendentesEscalasFolgas_=function(pm,permissao){
  const lista=FT_OC_PENDENTES_ANTERIOR_(pm,permissao), vistos=new Set(lista.map(p=>ocRe_(p.re)));
  const pedidos=epLer_('FT_PEDIDOS_FOLGA_PEL'), hoje=fcHoje_(), concedidos=fcIndice_();
  lista.forEach(p=>{
    p.folgaFlagrante=ocResumoFF_(p.re).pendentes.length;
    const re=ocRe_(p.re);
    p.folgaCompensatoria=concedidos[re]?(p.pelotao==='ADM'?fcTotalAposBaixas_(pedidos,re):fcLivres_(pedidos,re,'',hoje).length):0;
  });
  if(Object.keys(concedidos).some(re=>!vistos.has(re)))ocEfetivo_().forEach(p=>{
    const re=ocRe_(p.re),pelotao=grupoFrequencia_(p),alvo=Object.assign({},p,{pelotao});
    if(vistos.has(re)||!concedidos[re]||!usuarioPodeVerPmEscalasFolgas_(alvo,pm,permissao))return;
    lista.push({re:p.re,postoGrad:p.postoGrad,nomeGuerra:p.nomeGuerra,pelotao,folgaMensal:0,diasTrabalhados:0,folgaFlagrante:0,
      folgaCompensatoria:pelotao==='ADM'?fcTotalAposBaixas_(pedidos,re):fcLivres_(pedidos,re,'',hoje).length});
  });
  return lista;
};

// Folga compensatória: concessões em livro próprio e baixa no pedido aprovado.
// O mesmo ScriptLock protege concessões, pedidos, baixas e Controle de Folgas.
const FT_FC_ABA_='FT_CONCESSOES_FOLGA_COMPENSATORIA';
let FT_FC_CONCESSOES_=null;
let FT_FC_BLOQUEIOS_OFICIAIS_=null;
function fcPodeGerir_(pm){return ocCap_(pm).importar;} // somente Enc Adm e Gestor
function fcHoje_(){return Utilities.formatDate(new Date(),'America/Sao_Paulo','yyyy-MM-dd');}
function fcIndice_(){
  if(FT_FC_CONCESSOES_===null){
    FT_FC_CONCESSOES_={};
    epLer_(FT_FC_ABA_).forEach(function(c) {
      if (!FT_FC_CONCESSOES_[c.re]) FT_FC_CONCESSOES_[c.re] = [];
      FT_FC_CONCESSOES_[c.re].push(c);
    });
  }
  return FT_FC_CONCESSOES_;
}
function fcCreditos_(re){
  return (fcIndice_()[re]||[]).slice().sort((a,b)=>a.criadaEm.localeCompare(b.criadaEm)||a.id.localeCompare(b.id))
    .flatMap(c=>Array.from({length:c.quantidade},(_,i)=>({id:c.id+':'+(i+1),data:c.data})));
}
function fcLivres_(pedidos,re,data,hoje,ignorarId){
  const ocupados=new Set(pedidos.filter(p=>p.id!==ignorarId&&p.re===re&&p.tipo==='FC'&&
    (['APROVADA','GOZADA'].includes(p.estado)||(p.estado==='PENDENTE'&&p.data>=hoje))).map(p=>p.credito));
  return fcCreditos_(re).filter(c=>(!data||c.data<=data)&&!ocupados.has(c.id));
}
function fcTotalAposBaixas_(lista,re){
  const utilizados=new Set(lista.filter(p=>p.re===re&&p.tipo==='FC'&&
    (p.estado==='GOZADA'||p.pel==='ADM'&&p.estado==='APROVADA')).map(p=>p.credito));
  return fcCreditos_(re).filter(c=>!utilizados.has(c.id)).length;
}
function fcConflitoOficial_(re,data){
  return listarLancamentosAtivosFolgas_(re).some(x=>x.dataISO===data&&x.codigo&&x.codigo!=='FOL');
}
function fcBloqueioOficial_(re,data){
  if(FT_FC_BLOQUEIOS_OFICIAIS_===null)FT_FC_BLOQUEIOS_OFICIAIS_=new Set(epLer_('FT_PEDIDOS_FOLGA_PEL')
    .filter(p=>p.tipo==='FC'&&['APROVADA','GOZADA'].includes(p.estado)).map(p=>p.re+'|'+p.data));
  return FT_FC_BLOQUEIOS_OFICIAIS_.has(re+'|'+data);
}
function fcExecutar_(acao,params,auth){
  if(!fcPodeGerir_(auth.pm))throw new Error('A gestão de FC é restrita ao Enc Adm e ao Gestor.');
  if(acao==='folga_comp_gestao'){
    const efetivo=ocEfetivo_().map(ocPmSeguro_),hoje=fcHoje_();
    const pedidos=epLer_('FT_PEDIDOS_FOLGA_PEL').filter(p=>p.tipo==='FC'&&p.pel!=='ADM'&&p.estado==='APROVADA'&&p.data<=hoje)
      .sort((a,b)=>a.data.localeCompare(b.data)||a.nome.localeCompare(b.nome))
      .map(({id,nome,pel,data})=>({id,nome,pel,data}));
    const nomes=new Map(efetivo.map(p=>[ocRe_(p.re),[p.posto,p.nomeGuerra||p.nome].filter(Boolean).join(' ')]));
    const concessoes=Object.values(fcIndice_()).flat().sort((a,b)=>b.criadaEm.localeCompare(a.criadaEm))
      .slice(0,20).map(({re,data,quantidade,motivo,autor})=>({nome:nomes.get(re)||re,data,quantidade,motivo,autor}));
    return {efetivo,pedidos,concessoes};
  }
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const reAtor=ocRe_(auth.pm.re), ator=[auth.pm.postoGrad,auth.pm.nomeGuerra||auth.pm.nomeCompleto,auth.pm.re].filter(Boolean).join(' ');
    const agora=new Date().toISOString(),hoje=fcHoje_();
    if(acao==='folga_comp_conceder'){
      const id=ocTexto_(params.id),re=ocRe_(params.re),quantidade=Number(params.quantidade),motivo=ocTexto_(params.motivo);
      if(!/^[a-z0-9-]{16,64}$/i.test(id))throw new Error('Identificador de lançamento inválido.');
      if(!Number.isInteger(quantidade)||quantidade<1||quantidade>365)throw new Error('Informe uma quantidade inteira de 1 a 365 dias.');
      if(!motivo||motivo.length>400)throw new Error('Informe o motivo da concessão (até 400 caracteres).');
      if(!re||!ocEfetivo_().some(p=>ocRe_(p.re)===re))throw new Error('Selecione um policial do efetivo atual.');
      const anterior=Object.values(fcIndice_()).flat().find(c=>c.id===id);
      if(anterior){
        if(anterior.re!==re||anterior.autorRE!==reAtor||anterior.quantidade!==quantidade||anterior.motivo!==motivo)throw new Error('Identificador já utilizado. Atualize a tela.');
        return {id,quantidade};
      }
      epGravar_(FT_FC_ABA_,{id,re,quantidade,motivo,data:hoje,criadaEm:agora,autorRE:reAtor,autor:ator});
      FT_FC_CONCESSOES_=null;
      return {id,quantidade};
    }
    if(['folga_comp_gozo','folga_comp_devolver'].includes(acao)){
      const pedidos=epLer_('FT_PEDIDOS_FOLGA_PEL'),p=pedidos.find(p=>p.id===params.id&&p.tipo==='FC');
      const devolucao=acao==='folga_comp_devolver',motivo=ocTexto_(params.motivo);
      if(p&&p.pel==='ADM')throw new Error('A FC da ADM é baixada na aprovação em Escalas ADM.');
      if((p == null ? void 0 : p.estado)===(devolucao?'CANCELADA':'GOZADA'))return {id:p.id};
      if(devolucao&&(!motivo||motivo.length>160))throw new Error('Informe o motivo da devolução (até 160 caracteres).');
      if(!p||p.estado!=='APROVADA'||p.data>hoje)throw new Error('Só é possível dar baixa em FC aprovada com data de gozo até hoje.');
      if(devolucao){
        p.estado='CANCELADA';p.motivo=motivo;p.devolvidaEm=agora;p.devolvidaPor=ator;p.devolvidaPorRE=reAtor;
      }else{
        if(fcConflitoOficial_(p.re,p.data))throw new Error('Há outro lançamento no Controle de Folgas para essa data. Confira antes de dar baixa.');
        if(!fcLivres_(pedidos,p.re,p.data,hoje,p.id).some(c=>c.id===p.credito))throw new Error('O crédito reservado não está disponível. Confira a concessão.');
        p.estado='GOZADA';p.gozoRegistradoEm=agora;p.gozoRegistradoPor=ator;p.gozoRegistradoPorRE=reAtor;
      }
      epGravar_('FT_PEDIDOS_FOLGA_PEL',p);
      FT_FC_BLOQUEIOS_OFICIAIS_=null;
      return {id:p.id};
    }
    throw new Error('Ação desconhecida.');
  }finally{lock.releaseLock();}
}

// Espaço do Pel. Os pedidos não dão baixa ao serem aprovados.
function epData_(d,n){return new Date(Date.parse(d+'T00:00:00Z')+n*86400000).toISOString().slice(0,10);}
function epSemana_(d){return epData_(d,-((new Date(d+'T00:00:00Z').getUTCDay()+5)%7));}
function epGrupo_(p){
  const grupo=normalizarPelotaoResenha_(p.pelotao||'');
  return ['1º PEL','2º PEL','ADM'].includes(grupo)?grupo:grupoFrequencia_(p);
}
const FT_EP_INICIO_12X36_='2026-11-16';
function epDiaPadrao_(d,pel){
  if(!['1º PEL','2º PEL'].includes(pel))return false;
  if(d>=FT_EP_INICIO_12X36_){
    const dias=(Date.parse(d+'T00:00:00Z')-Date.parse(FT_EP_INICIO_12X36_+'T00:00:00Z'))/86400000;
    return pel===(dias%2===0?'2º PEL':'1º PEL');
  }
  const n=Math.round((Date.parse(d+'T00:00:00Z')-Date.parse('2026-11-02T00:00:00Z'))/86400000);
  const w=((Math.floor(n/7)%2)+2)%2, dia=((n%7)+7)%7;
  return (dia===1||dia===3?1-w:w)===(pel==='2º PEL'?0:1);
}
function epDiaEfetivo_(data,pel,config){
  const padrao=epDiaPadrao_(data,pel);
  if(pel==='ADM'&&config)return config;
  if(config&&(data<FT_EP_INICIO_12X36_||config.operacao||config.regime12x36===true))return config;
  // Não herdar marcações de serviço do regime anterior. Mantém as vagas
  // personalizadas onde o dia já era de serviço e continua compatível.
  return {data,pel,linha:(config == null ? void 0 : config.linha),servico:padrao,operacao:false,
    vagas:padrao&&(config == null ? void 0 : config.servico)?config.vagas:1};
}
function epLer_(nome){
  const sh=obterAbaSeExistir_(nome);
  return !sh||sh.getLastRow()<2?[]:sh.getRange(2,1,sh.getLastRow()-1,1).getValues()
    .map((r,i)=>r[0]?Object.assign(JSON.parse(r[0]),{linha:i+2}):null).filter(Boolean);
}
function epGravar_(nome,r){
  const sh=obterOuCriarAbaComCabecalho_(nome,['REGISTRO_JSON']);
  const copia=Object.assign({},r);delete copia.linha;
  sh.getRange(r.linha||sh.getLastRow()+1,1).setValue(JSON.stringify(copia));
}
function epSaldo_(re){
  const fm=obterSaldosAtuaisFolgasPendentes_()[re];
  const saldo=fm||reconstruirEstadoFolgas_(re,obterSaldosIniciaisFolgas_(),listarLancamentosAtivosFolgas_(re));
  if(!fm&&!saldo.ok)throw new Error('Não foi possível conferir o saldo de FM.');
  return {fm:Math.max(0,Number(saldo.saldoFolgas)||0),ff:ocResumoFF_(re).pendentes||[]};
}
function epReservas_(lista,re,hoje,ignorarId){
  const oficial={};
  listarLancamentosAtivosFolgas_(re).forEach(e=>{oficial[e.dataISO]=e.codigo;});
  return lista.concat(epReservasUrgentes_(re,hoje,ignorarId)).filter(p=>p.re===re&&['PENDENTE','APROVADA'].includes(p.estado)&&p.data>=hoje&&oficial[p.data]!==p.tipo);
}
function epSaldoDisponivel_(lista,re,hoje,adm){
  const FC=fcLivres_(lista,re,'',hoje).length;
  if(adm)return {FM:0,FF:0,FC,FCTotal:fcTotalAposBaixas_(lista,re)};
  const saldo=epSaldo_(re),reservas=epReservas_(lista,re,hoje);
  return {FM:Math.max(0,saldo.fm-reservas.filter(x=>x.tipo==='FM').length),
    FF:Math.max(0,saldo.ff.length-reservas.filter(x=>x.tipo==='FF').length),FC};
}
function epExecutar_(acao,params,auth,viaAdm){
  const pm=auth.pm,perfil=ocPerfil_(pm),gestor=auth.cap.ffDelegada===true,encAdm=auth.cap.importar===true;
  const proprio=epGrupo_(pm);
  const solicitado=normalizarTexto_(params.pel||proprio),pel=solicitado==='ADM'?'ADM':normalizarPelotaoResenha_(params.pel||proprio);
  const adm=pel==='ADM',gerirAdm=fcPodeGerir_(pm);
  if(adm&&viaAdm!==true&&!['espaco_pel_mensagens','espaco_pel_mensagens_contador','espaco_pel_responder_liberacao'].includes(acao))throw new Error('Solicite ou gerencie FC da ADM em Escalas ADM.');
  if(viaAdm===true&&!adm)throw new Error('Esta agenda é exclusiva da ADM.');
  const sgt=/SGT|SARGENTO/.test(normalizarTexto_(pm.postoGrad||''))||['CMT_PEL','CMT_PELOTAO','COMANDANTE_PELOTAO','COMANDANTE_DE_PELOTAO'].includes(perfil);
  const gerencia=adm?gerirAdm:gestor||encAdm||(sgt&&pel===proprio);
  if(!['1º PEL','2º PEL','ADM'].includes(pel)||(pel!==proprio&&!(adm?gerirAdm:gestor||encAdm)))throw new Error('Acesso restrito à sua lotação.');
  const hoje=Utilities.formatDate(new Date(),'America/Sao_Paulo','yyyy-MM-dd');
  const agora=Utilities.formatDate(new Date(),'America/Sao_Paulo',"yyyy-MM-dd'T'HH:mm");
  const data=String(params.data||epData_(epSemana_(hoje),7));
  if(!ocISOValida_(data))throw new Error('Data inválida.');
  const semana=epSemana_(data),fim=epData_(semana,6),limite=epData_(semana,-3)+'T10:00';
  const re=ocRe_(pm.re),nome=[pm.postoGrad,pm.nomeGuerra||pm.nomeCompleto].filter(Boolean).join(' ');
  const pedidos='FT_PEDIDOS_FOLGA_PEL',cal='FT_DIAS_ESPACO_PEL',liberacoes='FT_LIBERACOES_FOLGA_PEL',solicitacoes='FT_SOLICITACOES_ABERTURA_FOLGA';
  const leitura=['espaco_pel_dados','espaco_pel_resenha','espaco_pel_saldo','espaco_pel_mensagens','espaco_pel_mensagens_contador','espaco_pel_destinatarios'].includes(acao);
  const lock=leitura?null:LockService.getScriptLock();
  if(lock)lock.waitLock(30000);
  try{
    if(acao==='espaco_pel_saldo'&&pel!==proprio)throw new Error('Consulta restrita ao próprio saldo.');
    if(acao==='espaco_pel_resenha'&&(adm||!gerencia&&!perfilCmtEquipeResenha_(pm)))throw new Error('Consulta restrita ao responsável pela resenha do Pel ou ao Gestor.');
    if(acao==='espaco_pel_mensagens_contador'){const cache=CacheService.getScriptCache(),key='EP_MSG_COUNT_'+re,cached=cache.get(key);if(cached!==null)return {total:Number(cached)||0};const msgs=epLer_(solicitacoes),total=msgs.filter(x=>x.destinatario===re&&x.estado==='PENDENTE').length;cache.put(key,String(total),60);return {total};}
    if(acao==='espaco_pel_mensagens'){const msgs=epLer_(solicitacoes);return {itens:msgs.filter(x=>x.destinatario===re||x.solicitante===re).sort((a,b)=>String(b.criadoEm).localeCompare(String(a.criadoEm))).map(x=>({id:x.id,data:x.data,pel:x.pel,tipo:x.tipo,motivo:x.motivo,estado:x.estado,solicitante:x.solicitanteNome,destinatario:x.destinatarioNome,recebido:x.destinatario===re,decisor:x.decisorNome||''}))};}
    const todos=epLer_(pedidos),grants=epLer_(liberacoes),msgs=epLer_(solicitacoes);
    if(acao==='espaco_pel_saldo')return epSaldoDisponivel_(todos,re,hoje,adm);
    const cfg=new Map();
    epLer_(cal).forEach(x=>{if(x.pel===pel&&x.data>=semana&&x.data<=fim&&!cfg.has(x.data))cfg.set(x.data,x);});
    const dia=d=>epDiaEfetivo_(d,pel,cfg.get(d));
    const grantFor=(d,tipo)=>grants.find(x=>x.re===re&&x.pel===pel&&x.data===d&&(!tipo||x.tipo===tipo)&&!x.usada);
    const diaAberto=(d,tipo)=>agora<limite||!!(cfg.get(d)&&cfg.get(d).aberturaManual)||!!grantFor(d,tipo);
    const lista=todos.filter(x=>x.pel===pel&&x.data>=semana&&x.data<=fim);
    if(acao==='espaco_pel_resenha'){
      // A ausência já aprovada precisa aparecer mesmo se o dia tiver sido
      // reconfigurado como operação/folga. Escondê-la aqui liberava o policial.
      const itens=lista.filter(x=>x.data===data&&(x.estado==='APROVADA'||x.tipo==='FC'&&x.estado==='GOZADA')).map(x=>({id:x.id,re:x.re,nome:x.nome,tipo:x.tipo,bopm:x.bopm||'',referencia:x.referencia||''}));
      // Pedido de FF ainda pendente não deve ser lançado como ausência aprovada,
      // mas também não pode deixar o policial selecionável na equipe.
      const bloqueiosEquipe=lista.filter(x=>x.data===data&&['FF','FM','FC'].includes(x.tipo)&&x.estado==='PENDENTE')
        .map(x=>({re:x.re,nome:x.nome,tipo:x.tipo,estado:x.estado}));
      return {itens,bloqueiosEquipe};
    }
    if(acao==='espaco_pel_dados'){
      const saldo=params.semSaldo==='1'?null:proprio===pel?epSaldoDisponivel_(todos,re,hoje,adm):{FM:0,FF:0,FC:0};
      const confirmadas={};
      lista.forEach(p=>{if(p.estado==='APROVADA'||p.tipo==='FC'&&p.estado==='GOZADA')confirmadas[p.data]=(confirmadas[p.data]||0)+1;});
      return {pel,proprio,gestor,podeAlternar:gestor||encAdm,gerencia,semana,fim,limite,aberto:agora<limite,saldo,
        dias:Array.from({length:7},(_,i)=>{const d=dia(epData_(semana,i));return {data:d.data,servico:d.servico,operacao:d.operacao,vagas:d.vagas,
          confirmadas:confirmadas[d.data]||0,aberturaManual:!!(cfg.get(d.data)&&cfg.get(d.data).aberturaManual),liberado:!!grantFor(d.data),liberadoTipo:(grantFor(d.data)||{}).tipo||''};}),
        pedidos:lista.filter(p=>gerencia||p.re===re).map(p=>({id:p.id,re:p.re,nome:p.nome,tipo:p.tipo,data:p.data,estado:p.estado,motivo:p.motivo||'',meu:p.re===re,foraEscala:p.data>=FT_EP_INICIO_12X36_&&!dia(p.data).servico})),
        pedidosColegas:adm||gerencia?[]:lista.filter(p=>p.re!==re&&['PENDENTE','APROVADA','GOZADA'].includes(p.estado))
          .sort((a,b)=>a.data.localeCompare(b.data)||a.nome.localeCompare(b.nome))
          .map(p=>({nome:p.nome,data:p.data,estado:p.estado}))};
    }
    if(acao==='espaco_pel_liberar_data'){
      if(!gerencia||data<hoje||agora<limite)throw new Error('A liberação manual só pode ser feita por responsável após o prazo.');
      const anterior=dia(data),aberta=String(params.aberta)==='true';
      if(!anterior.servico||anterior.operacao)throw new Error('Só é possível abrir pedidos em dia de serviço sem operação.');
      const existente=cfg.get(data)||anterior;
      epGravar_(cal,{linha:existente.linha,pel,data,servico:anterior.servico,operacao:anterior.operacao,vagas:anterior.vagas,regime12x36:data>=FT_EP_INICIO_12X36_,aberturaManual:aberta,autor:re,alterado:new Date().toISOString()});
      return {aberta};
    }
    if(acao==='espaco_pel_destinatarios'){
      if(proprio!==pel||adm)throw new Error('Solicitação disponível somente aos policiais dos pelotões.');
      const itens=ocEfetivo_().filter(x=>{const xpel=epGrupo_(x),xp=ocPerfil_(x),xs=/SGT|SARGENTO/.test(normalizarTexto_(x.postoGrad||''))||['CMT_PEL','CMT_PELOTAO','COMANDANTE_PELOTAO'].includes(xp);return x.ativo!==false&&(xpel===pel&&xs||auth.cap.importar&&['ENC_ADM','ENC_ADM_FT','ENCARREGADO_ADM_FT','ENCARREGADO_ADMINISTRATIVO'].includes(xp)||xp.indexOf('GESTOR')>=0);}).map(x=>({re:ocRe_(x.re),nome:[x.postoGrad,x.nomeGuerra||x.nomeCompleto].filter(Boolean).join(' '),pel:epGrupo_(x)})).filter(x=>x.re&&x.re!==re);
      return {itens};
    }
    if(acao==='espaco_pel_pedir_liberacao'){
      if(proprio!==pel||adm||agora<limite||data<hoje||!dia(data).servico||dia(data).operacao)throw new Error('A solicitação de liberação só está disponível após o prazo e para data de serviço futura.');
      const tipo=String(params.tipo||'');if(!['FM','FF','FC'].includes(tipo))throw new Error('Tipo de folga inválido.');
      const alvo=ocRe_(params.destinatario),dest=ocEfetivo_().find(x=>ocRe_(x.re)===alvo);if(!dest||alvo===re)throw new Error('Selecione um destinatário disponível.');
      const dp=ocPerfil_(dest),dPel=epGrupo_(dest),destSgt=/SGT|SARGENTO/.test(normalizarTexto_(dest.postoGrad||''))||['CMT_PEL','CMT_PELOTAO','COMANDANTE_PELOTAO'].includes(dp),destAdm=['ENC_ADM','ENC_ADM_FT','ENCARREGADO_ADM_FT','ENCARREGADO_ADMINISTRATIVO'].includes(dp),destGestor=dp.indexOf('GESTOR')>=0;
      if(!(destSgt&&dPel===pel||destAdm||destGestor)||msgs.some(x=>x.solicitante===re&&x.data===data&&x.estado==='PENDENTE'))throw new Error('Destinatário inválido ou já existe solicitação pendente para esta data.');
      const motivo=ocTexto_(params.motivo).slice(0,300);if(motivo.length<5)throw new Error('Informe a motivação da solicitação.');
      epGravar_(solicitacoes,{id:Utilities.getUuid(),solicitante:re,solicitanteNome:nome,destinatario:alvo,destinatarioNome:[dest.postoGrad,dest.nomeGuerra||dest.nomeCompleto].filter(Boolean).join(' '),pel,data,tipo,motivo,estado:'PENDENTE',criadoEm:new Date().toISOString()});CacheService.getScriptCache().remove('EP_MSG_COUNT_'+alvo);return {};
    }
    if(acao==='espaco_pel_responder_liberacao'){
      const m=msgs.find(x=>x.id===params.id&&x.destinatario===re&&x.estado==='PENDENTE');if(!m)throw new Error('Solicitação não disponível para este destinatário.');
      const decisao=String(params.decisao||'');if(!['AUTORIZADA','RECUSADA'].includes(decisao))throw new Error('Decisão inválida.');
      const atual=ocEfetivo_().find(x=>ocRe_(x.re)===re),perfilAtual=ocPerfil_(atual||{}),pelAtual=epGrupo_(atual||{}),sgtAtual=/SGT|SARGENTO/.test(normalizarTexto_((atual||{}).postoGrad||''))||['CMT_PEL','CMT_PELOTAO','COMANDANTE_PELOTAO'].includes(perfilAtual),encAtual=['ENC_ADM','ENC_ADM_FT','ENCARREGADO_ADM_FT','ENCARREGADO_ADMINISTRATIVO'].includes(perfilAtual),gestorAtual=perfilAtual.indexOf('GESTOR')>=0;
      if(!(sgtAtual&&pelAtual===m.pel||encAtual||gestorAtual))throw new Error('Seu perfil não está mais autorizado a decidir esta solicitação.');
      if(decisao==='AUTORIZADA'){
        const calAtual=epLer_(cal).find(x=>x.pel===m.pel&&x.data===m.data),diaPedido=epDiaEfetivo_(m.data,m.pel,calAtual);
        if(m.data<hoje||!diaPedido.servico||diaPedido.operacao)throw new Error('A data não pode mais receber solicitação.');
        epGravar_(liberacoes,{id:Utilities.getUuid(),re:m.solicitante,pel:m.pel,data:m.data,tipo:m.tipo,solicitacaoId:m.id,usada:false,autor:re,criadoEm:new Date().toISOString()});
      }
      m.estado=decisao;m.decisor=re;m.decisorNome=nome;m.respondidoEm=new Date().toISOString();epGravar_(solicitacoes,m);CacheService.getScriptCache().remove('EP_MSG_COUNT_'+re);return {};
    }
    if(acao==='espaco_pel_configurar'){
      if(!gerencia||data<hoje)throw new Error('Sem permissão para alterar este dia.');
      const anterior=dia(data);
      const operacao=String(params.operacao)==='true',servico=operacao||String(params.servico)==='true',vagas=operacao||!servico?0:Number(params.vagas);
      if(!Number.isInteger(vagas)||vagas<0||vagas>50)throw new Error('Informe de 0 a 50 vagas.');
      const afetados=lista.filter(p=>p.data===data&&['PENDENTE','APROVADA'].includes(p.estado));
      if(servico&&!operacao&&afetados.filter(p=>p.estado==='APROVADA').length>vagas)throw new Error('Há mais folgas aprovadas do que o novo limite.');
      if((operacao||!servico)&&afetados.length&&params.confirmar!=='SIM')throw new Error('Confirme o cancelamento dos pedidos deste dia.');
      epGravar_(cal,{linha:anterior.linha,pel,data,servico,operacao,vagas,regime12x36:data>=FT_EP_INICIO_12X36_,aberturaManual:!!anterior.aberturaManual,autor:re,alterado:new Date().toISOString()});
      if(operacao||!servico)afetados.forEach(p=>{p.estado='CANCELADA';p.motivo=operacao?'Cancelada por operação':'Dia retirado da escala';p.decisor=re;p.decididaEm=new Date().toISOString();epGravar_(pedidos,p);});
      return {};
    }
    if(acao==='espaco_pel_pedir'){
      const tipo=String(params.tipo||''),grant=grantFor(data,tipo),d=dia(data);if(proprio!==pel||!diaAberto(data,tipo)||data<hoje)throw new Error('Prazo encerrado ou pelotão diferente do cadastro.');
      if(!(adm?['FC']:['FM','FF','FC']).includes(tipo)||!d.servico||d.operacao||d.vagas<1)throw new Error('Este dia não está disponível para pedidos.');
      if(tipo==='FC'?fcConflitoOficial_(re,data):listarLancamentosAtivosFolgas_(re).some(x=>x.dataISO===data&&['FF','FM','DS'].includes(x.codigo)))throw new Error('Já existe outro lançamento para este dia.');
      if(todos.some(x=>x.re===re&&x.data===data&&['PENDENTE','APROVADA','GOZADA'].includes(x.estado)))throw new Error('Você já tem pedido para este dia.');
      const reservas=epReservas_(todos,re,hoje);
      if(reservas.some(x=>x.data===data))throw new Error('Já existe uma folga reservada para este dia.');
      let c=null;
      if(tipo==='FC'){
        c=fcLivres_(todos,re,data,hoje)[0];
        if(!c)throw new Error('Não possui Folga Compensatória disponível para fruição nesta data.');
      }else{
        const saldo=epSaldo_(re),livres=saldo.ff.filter(c=>c.data<=data&&!reservas.some(x=>x.credito===c.id));
        if(tipo==='FM'&&saldo.fm<=reservas.filter(x=>x.tipo==='FM').length||tipo==='FF'&&!livres.length)throw new Error('Não possui Folga '+(tipo==='FM'?'Mensal':'Flagrante')+' para fruição.');
        if(tipo==='FF')c=livres[0];
      }
      const novoId=Utilities.getUuid();if(grant){grant.usada=true;grant.pedidoId=novoId;epGravar_(liberacoes,grant);}
      epGravar_(pedidos,{id:novoId,re,nome,pel,data,tipo,estado:'PENDENTE',credito:(c == null ? void 0 : c.id)||'',bopm:(c == null ? void 0 : c.bopm)||'',referencia:(c == null ? void 0 : c.data)||'',enviadoEm:new Date().toISOString()});
      return {};
    }
    if(acao==='espaco_pel_decidir'){
      const p=lista.find(x=>x.id===params.id&&x.data===data);
      if(!p||!['PENDENTE','APROVADA'].includes(p.estado))throw new Error('Pedido já decidido ou indisponível.');
      const decisao=String(params.decisao||'');
      const cancelamento=decisao==='CANCELADA'&&p.re===re&&p.estado==='PENDENTE'&&agora<limite;
      const cancelarAprovada=decisao==='CANCELADA'&&p.estado==='APROVADA'&&gerencia&&p.re!==re&&data>=hoje;
      if(cancelarAprovada&&!ocTexto_(params.motivo))throw new Error('Informe o motivo do cancelamento.');
      if(!cancelamento&&!cancelarAprovada&&(p.estado!=='PENDENTE'||!gerencia||p.re===re||!['APROVADA','RECUSADA'].includes(decisao)))throw new Error('Sem permissão para decidir este pedido.');
      if(decisao==='APROVADA'){
        const d=dia(p.data);
        if(p.tipo==='FC'?fcConflitoOficial_(p.re,p.data):listarLancamentosAtivosFolgas_(p.re).some(x=>x.dataISO===p.data&&['FF','FM','DS'].includes(x.codigo)))throw new Error('Já existe outro lançamento para este dia.');
        if(p.data<hoje||!d.servico||d.operacao||lista.filter(x=>x.data===p.data&&['APROVADA','GOZADA'].includes(x.estado)).length>=d.vagas)throw new Error('Não há vaga para aprovar neste dia.');
        if(p.tipo==='FC'){
          if(epReservasUrgentes_(p.re,hoje).some(x=>x.data===p.data))throw new Error('Já existe folga de urgência reservada na resenha para esse dia.');
          if(!fcLivres_(todos,p.re,p.data,hoje,p.id).some(c=>c.id===p.credito))throw new Error('Crédito de FC indisponível no momento da aprovação.');
        }else{
          const saldo=epSaldo_(p.re),outras=epReservas_(todos,p.re,hoje).filter(x=>x.id!==p.id);
          if(p.tipo==='FM'&&saldo.fm<=outras.filter(x=>x.tipo==='FM').length||p.tipo==='FF'&&!saldo.ff.some(c=>c.id===p.credito&&c.data<=p.data))throw new Error('Saldo indisponível no momento da aprovação.');
        }
      }
      p.estado=decisao;p.motivo=ocTexto_(params.motivo).slice(0,160);p.decisor=re;p.decididaEm=new Date().toISOString();
      // A decisão e a baixa da ADM são gravadas juntas no próprio pedido.
      if(adm&&decisao==='APROVADA')p.baixaAutomaticaEm=p.decididaEm;
      if(adm&&decisao==='CANCELADA'&&cancelarAprovada)p.estornoAutomaticoEm=p.decididaEm;
      epGravar_(pedidos,p);
      return {baixaAutomatica:adm&&decisao==='APROVADA'};
    }
    throw new Error('Ação desconhecida.');
  }finally{if(lock)lock.releaseLock();}
}


// Reservas urgentes persistem no próprio payload da resenha. A leitura e a
// gravação ficam sob o mesmo ScriptLock usado pelo fluxo de escalas.
function epReservasUrgentes_(re,hoje,ignorarId){
  const abas=[obterAbaSeExistir_(FT_API.abas.resenhasApp),obterAbaSeExistir_('RESENHAS_ESCALAS')].filter((a,i,l)=>a&&l.indexOf(a)===i);
  const vistos={},ret=[];
  for(const aba of abas){
    if(aba.getLastRow()<2)continue;
    for(const linha of aba.getRange(2,1,aba.getLastRow()-1,15).getValues()){
      const id=String(linha[0]||'');
      if(!id||id===ignorarId||vistos[id])continue;
      vistos[id]=true;
      const data=dataBRParaISOApi_(linha[2]);
      if(!data||data<hoje||[linha[7],linha[13]].some(v=>/EXCLUID|CANCELAD|RECUSAD/.test(normalizarTexto_(v||''))))continue;
      const p=parseJsonParam_(linha[9],{});
      (Array.isArray(p.urgenciasPel)?p.urgenciasPel:[]).forEach(x=>{if(x.re===re)ret.push({re,tipo:x.tipo,data,credito:x.credito||'',estado:'APROVADA'});});
    }
  }
  return ret;
}
function epCodigoResenha_(valor){
  const s=normalizarTexto_(valor||'');
  return ['FM','FOLGA MENSAL'].includes(s)?'FM':['FF','FOLGA FLAGRANTE'].includes(s)?'FF':['FC','FOLGA COMPENSATORIA'].includes(s)?'FC':/DISPENSA.*SERVICO|^DS$/.test(s)?'DS':'';
}
function epValidarFolgasResenha_(payload,pm,id,anterior){
  // Somente o servidor pode criar os metadados que geram reserva.
  delete payload.urgenciasPel;
  if(payload.itens==null)payload.itens=[];
  if(!Array.isArray(payload.itens))throw new Error('Itens da resenha inválidos.');
  const itens=payload.itens;
  for(let i=itens.length-1;i>=0;i--){
    if((itens[i] == null ? void 0 : itens[i].tipo)==='FOLGA_DISPENSA'&&normalizarTexto_(itens[i].campo||'')==='REGISTRO DA URGENCIA')itens.splice(i,1);
  }
  const grupos=new Map();
  for(const i of itens){
    if(i.tipo!=='FOLGA_DISPENSA')continue;
    const k=String(i.grupo),g=grupos.get(k)||{};
    const nome=normalizarTexto_(i.campo||'');
    if(g[nome])throw new Error('Campo de Folga/Dispensa duplicado.');
    g[nome]=i;grupos.set(k,g);
  }
  let temFolga=Array.from(grupos.values()).some(g=>['FM','FF','FC'].includes(epCodigoResenha_((g.TIPO == null ? void 0 : g.TIPO.valorTecnico)||(g.TIPO == null ? void 0 : g.TIPO.valor))));
  const pel=normalizarPelotaoResenha_(payload.pelotao),data=String(payload.dataEscala||'');
  if(!ocISOValida_(data)||!['1º PEL','2º PEL'].includes(pel)){
    if(temFolga)throw new Error('Data ou pelotão inválido para a folga.');
    return;
  }
  const pedidos=epLer_('FT_PEDIDOS_FOLGA_PEL');
  // Reconstitui no servidor as FF/FM aprovadas. A inclusão automática não pode
  // depender apenas do navegador ter carregado ou mesclado a prévia.
  const aprovadasAuto=pedidos.filter(x=>x.pel===pel&&x.data===data&&['FF','FM'].includes(x.tipo)&&x.estado==='APROVADA');
  let grupoAuto=Math.max(0,...Array.from(grupos.keys()).map(Number).filter(Number.isFinite));
  for(const pedido of aprovadasAuto){
    const existente=Array.from(grupos.values()).find(g=>g.POLICIAL&&ocRe_(g.POLICIAL.valorTecnico||'')===pedido.re);
    if(existente){
      if(epCodigoResenha_(existente.TIPO&&(existente.TIPO.valorTecnico||existente.TIPO.valor))!==pedido.tipo)
        throw new Error((pedido.nome||pedido.re)+' já possui outro tipo de ausência nesta data.');
      continue;
    }
    if(itens.some(i=>['FT','ROCAM'].includes(i.tipo)&&ocRe_(i.valorTecnico||'')===pedido.re))
      throw new Error((pedido.nome||pedido.re)+' possui folga aprovada nesta data e não pode compor equipe.');
    const grupo=String(++grupoAuto),nome=pedido.nome||pedido.re;
    const novo=(campo,valor,valorTecnico=valor)=>({tipo:'FOLGA_DISPENSA',grupo,campo,valor,valorTecnico,origemPelAuto:true});
    const tipoRotulo=pedido.tipo==='FF'?'Folga Flagrante':'Folga Mensal';
    const policial=novo('Policial',nome+' — RE '+pedido.re,pedido.re),tipo=novo('Tipo',tipoRotulo,pedido.tipo);
    itens.push(policial,tipo);grupos.set(grupo,{POLICIAL:policial,TIPO:tipo});
    if(pedido.tipo==='FF'){
      const referencia=novo('Data de referência',pedido.referencia||'',pedido.referencia||'');
      const bopm=novo('BOPM',pedido.bopm||'',pedido.bopm||'');
      itens.push(referencia,bopm);grupos.get(grupo)['DATA DE REFERENCIA']=referencia;grupos.get(grupo).BOPM=bopm;
    }
  }
  temFolga=Array.from(grupos.values()).some(g=>['FM','FF','FC'].includes(epCodigoResenha_((g.TIPO == null ? void 0 : g.TIPO.valorTecnico)||(g.TIPO == null ? void 0 : g.TIPO.valor))));
  const folgasPendentes=pedidos.filter(x=>x.pel===pel&&x.data===data&&['FF','FM','FC'].includes(x.tipo)&&x.estado==='PENDENTE');
  const equipesPayload=itens.filter(i=>['FT','ROCAM'].includes(i.tipo));
  for(const pedido of folgasPendentes){
    if(equipesPayload.some(i=>ocRe_(i.valorTecnico||'')===pedido.re))
      throw new Error((pedido.nome||pedido.re)+' possui solicitação de folga pendente nesta data. Decida o pedido antes de compor a equipe.');
  }
  const fcs=pedidos.filter(x=>x.pel===pel&&x.data===data&&x.tipo==='FC'&&['APROVADA','GOZADA'].includes(x.estado));
  if(!temFolga&&!fcs.length)return;
  const hoje=Utilities.formatDate(new Date(),'America/Sao_Paulo','yyyy-MM-dd');
  const agora=Utilities.formatDate(new Date(),'America/Sao_Paulo',"yyyy-MM-dd'T'HH:mm");
  const limite=epData_(epSemana_(data),-3)+'T10:00';
  const antigos=anterior?parseJsonParam_(anterior[9],{}):{};
  const dataAnterior=anterior?dataBRParaISOApi_(anterior[2]):'';
  const estadoAnterior=anterior?normalizarTexto_(anterior[13]||anterior[7]||'').replace(/_/g,' '):'';
  const calend=epDiaEfetivo_(data,pel,epLer_('FT_DIAS_ESPACO_PEL').find(d=>d.pel===pel&&d.data===data));
  if(!calend.operacao&&calend.servico){
    let n=Math.max(0,...Array.from(grupos.keys()).map(Number).filter(Number.isFinite));
    for(const p of fcs){
      const atual=Array.from(grupos.values()).find(g=>g.TIPO&&ocRe_((g.POLICIAL == null ? void 0 : g.POLICIAL.valorTecnico)||'')===p.re);
      if(atual){
        if(epCodigoResenha_(atual.TIPO.valorTecnico||atual.TIPO.valor)!=='FC')throw new Error(p.nome+' já possui uma FC aprovada nessa data.');
        continue;
      }
      const grupo=String(++n);
      const novo=(campo,valor,valorTecnico=valor)=>({tipo:'FOLGA_DISPENSA',grupo,campo,valor,valorTecnico,origemPelAuto:true});
      const g={POLICIAL:novo('Policial',p.nome,p.re),TIPO:novo('Tipo','Folga Compensatória')};
      itens.push(g.POLICIAL,g.TIPO);
      grupos.set(grupo,g);
    }
  }
  const roster=ocEfetivo_(),vistos=new Set(),ausencias={};
  for(const g of grupos.values()){
    const re=ocRe_((g.POLICIAL == null ? void 0 : g.POLICIAL.valorTecnico)||'');
    if(re&&g.TIPO)ausencias[re]=(ausencias[re]||0)+1;
  }
  const urgencias=[];
  for(const [grupo,g] of grupos){
    const tipo=epCodigoResenha_((g.TIPO == null ? void 0 : g.TIPO.valorTecnico)||(g.TIPO == null ? void 0 : g.TIPO.valor));
    if(!['FM','FF','FC'].includes(tipo))continue;
    const re=ocRe_((g.POLICIAL == null ? void 0 : g.POLICIAL.valorTecnico)||''),alvo=roster.find(x=>ocRe_(x.re)===re&&epGrupo_(x)===pel);
    if(!re||!alvo)throw new Error('Selecione um policial ativo do pelotão para a folga.');
    if(vistos.has(re)||ausencias[re]>1)throw new Error('Há mais de uma folga ou dispensa para o mesmo policial.');
    vistos.add(re);
    if(itens.some(i=>['FT','ROCAM'].includes(i.tipo)&&ocRe_(i.valorTecnico||'')===re))throw new Error('O policial '+re+' está escalado e também consta em folga.');
    const fcRegistrada=fcs.find(x=>x.re===re);
    if(tipo==='FC'){
      if(!fcRegistrada){
        // Permite manter FC manual já existente em resenha anterior ao módulo.
        const anteriores=Array.isArray(antigos.itens)?antigos.itens:[];
        const legado=dataAnterior===data&&!pedidos.some(x=>x.re===re&&x.data===data&&x.tipo==='FC')&&anteriores.some(i=>
          i.tipo==='FOLGA_DISPENSA'&&normalizarTexto_(i.campo)==='POLICIAL'&&ocRe_(i.valorTecnico||'')===re&&anteriores.some(t=>
            t.tipo===i.tipo&&String(t.grupo)===String(i.grupo)&&normalizarTexto_(t.campo)==='TIPO'&&epCodigoResenha_(t.valorTecnico||t.valor)==='FC'));
        if(legado&&ocISOValida_((g['DATA DE REFERENCIA'] == null ? void 0 : g['DATA DE REFERENCIA'].valorTecnico)||(g['DATA DE REFERENCIA'] == null ? void 0 : g['DATA DE REFERENCIA'].valor)))continue;
        throw new Error('Folga Compensatória exige pedido aprovado no Espaço do Pelotão.');
      }
      if(calend.operacao||!calend.servico)throw new Error('Não é possível lançar FC fora do serviço regular.');
      if(fcConflitoOficial_(re,data))throw new Error(alvo.nomeGuerra+' tem outro lançamento no Controle de Folgas nessa data.');
      g.POLICIAL.valor=[alvo.postoGrad,alvo.nomeGuerra||alvo.nomeCompleto].filter(Boolean).join(' ')+' — RE '+alvo.re;
      g.POLICIAL.valorTecnico=alvo.re;
      const ref=fcRegistrada.referencia;
      const campo=g['DATA DE REFERENCIA'];
      if(campo){campo.valor=ref;campo.valorTecnico=ref;}
      else itens.push({tipo:'FOLGA_DISPENSA',grupo,campo:'Data de referência',valor:ref,valorTecnico:ref,origemPelAuto:!!g.POLICIAL.origemPelAuto});
      continue;
    }
    if(fcRegistrada)throw new Error('Já existe FC para '+alvo.nomeGuerra+' nessa data.');
    if(calend.operacao||!calend.servico)throw new Error('Não é possível lançar FF/FM em dia sem serviço regular ou marcado como operação.');
    const oficial=listarLancamentosAtivosFolgas_(re).find(e=>e.dataISO===data&&['FM','FF','DS'].includes(e.codigo));
    if(oficial&&oficial.codigo!==tipo)throw new Error('Já existe outro tipo de ausência no Controle de Folgas para '+re+'.');
    const aprovado=pedidos.find(x=>x.re===re&&x.pel===pel&&x.data===data&&x.tipo===tipo&&x.estado==='APROVADA');
    if(pedidos.some(x=>x.re===re&&x.data===data&&x.estado==='APROVADA'&&x.tipo!==tipo))throw new Error('Já há outro tipo de folga aprovada para '+re+'.');
    const previo=dataAnterior===data?(antigos.urgenciasPel||[]).find(x=>x.re===re&&x.tipo===tipo):null;
    // Sgt pode lançar FF/FM diretamente na resenha sem pedido prévio no Espaço do Pel.
    // A validação de saldo/crédito logo abaixo continua obrigatória no servidor.
    const nome=[alvo.postoGrad,alvo.nomeGuerra||alvo.nomeCompleto].filter(Boolean).join(' ');
    g.POLICIAL.valor=nome+' — RE '+alvo.re;g.POLICIAL.valorTecnico=alvo.re;
    const saldo=oficial?null:epSaldo_(re);
    const outras=oficial?[]:epReservas_(pedidos,re,hoje,id).filter(x=>!aprovado||x.id!==aprovado.id);
    if(!oficial&&outras.some(x=>x.data===data))throw new Error(nome+' já tem outra folga reservada para esta data.');
    if(!oficial&&tipo==='FM'&&saldo.fm<=outras.filter(x=>x.tipo==='FM').length)throw new Error(nome+' não possui Folga Mensal disponível para fruição.');
    let credito=null;
    if(tipo==='FF'){
      if(oficial){
        const c=ocResumoFF_(re).porDia[data];
        credito=c||aprovado&&{id:aprovado.credito,data:aprovado.referencia,bopm:aprovado.bopm}||previo&&{id:previo.credito,data:previo.referencia,bopm:previo.bopm};
        if(!credito)throw new Error('Não foi possível identificar o BOPM da FF já registrada para '+nome+'.');
      }else{
        const escolhido=(aprovado == null ? void 0 : aprovado.credito)||(previo == null ? void 0 : previo.credito);
        credito=saldo.ff.find(x=>x.data<=data&&(!escolhido||x.id===escolhido)&&!outras.some(y=>y.credito===x.id));
        if(!credito)throw new Error(nome+' não possui Folga Flagrante disponível para fruição nesta data.');
      }
    }
    function definir(campo,valor){
      const chave=normalizarTexto_(campo),i=g[chave];
      if(i){i.valor=valor;i.valorTecnico=valor;}
      else itens.push({tipo:'FOLGA_DISPENSA',grupo,campo,valor,valorTecnico:valor,origemPelAuto:!!g.POLICIAL.origemPelAuto});
    }
    if(credito){definir('Data de referência',credito.data);definir('BOPM',credito.bopm);}
    if(!aprovado&&!oficial||previo&&!aprovado){
      const justificativa=(g['MOTIVO DA URGENCIA'] == null ? void 0 : g['MOTIVO DA URGENCIA'].valorTecnico)||(g['MOTIVO DA URGENCIA'] == null ? void 0 : g['MOTIVO DA URGENCIA'].valor);
      const motivo=(previo == null ? void 0 : previo.motivo)||ocTexto_(justificativa)||'Incluída diretamente pelo Sgt na resenha.';
      if(agora>=limite&&motivo.length<5)throw new Error('Informe o motivo da inclusão após o prazo (mínimo 5 caracteres) para '+nome+'.');
      if(motivo.length>160)throw new Error('O motivo da inclusão deve ter até 160 caracteres.');
      const responsavel=(previo == null ? void 0 : previo.responsavel)||[pm.postoGrad,pm.nomeGuerra||pm.nomeCompleto,pm.re].filter(Boolean).join(' ');
      const quando=(previo == null ? void 0 : previo.quando)||new Date().toISOString();
      urgencias.push({re,tipo,motivo,responsavel,quando,credito:(credito == null ? void 0 : credito.id)||'',referencia:(credito == null ? void 0 : credito.data)||'',bopm:(credito == null ? void 0 : credito.bopm)||''});
      if(agora>=limite){
        definir('Motivo da urgência',motivo);
        definir('Registro da urgência','Lançada pelo Sgt após o prazo — '+responsavel+' — '+Utilities.formatDate(new Date(quando),'America/Sao_Paulo','dd/MM/yyyy HH:mm'));
      }
    }
  }
  if(urgencias.length)payload.urgenciasPel=urgencias;
}

/** Impede o segundo envio da mesma data. Em operação conjunta previamente
 * cadastrada para os dois Pelotões, permite uma resenha por Pelotão. */
function epVerificarBloqueioResenhaData_(dataEscala, pelotao) {
  const data = dataBRParaISOApi_(dataEscala);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(data)) return;

  const abas = [obterAbaSeExistir_(FT_API.abas.resenhasApp), obterAbaSeExistir_('RESENHAS_ESCALAS')];
  const idsAbas = new Set();
  const vistos = new Set();
  const existentes = [];
  abas.forEach(function(aba) {
    if (!aba || idsAbas.has(aba.getSheetId())) return;
    idsAbas.add(aba.getSheetId());
    const ultima = aba.getLastRow();
    if (ultima < 2) return;
    const largura = Math.min(8, aba.getLastColumn());
    if (largura < 8) return;
    aba.getRange(2, 1, ultima - 1, largura).getValues().forEach(function(linha) {
      const id = String(linha[0] || '').trim();
      if (!id || vistos.has(id) || dataBRParaISOApi_(linha[2]) !== data) return;
      vistos.add(id);
      const situacao = normalizarTexto_(linha[7] || '').replace(/\s+/g, '_');
      if (['EXCLUIDA_APP', 'EXCLUIDA', 'OCULTA_APP'].includes(situacao)) return;
      existentes.push(normalizarPelotaoResenha_(linha[3] || ''));
    });
  });
  if (!existentes.length) return;

  const operacao = (typeof epLer_ === 'function' ? epLer_('FT_OPERACOES_RESENHA') : [])
    .find(function(item) { return item.data === data && !item.canceladaEm && Array.isArray(item.pelotoes); });
  const pelotoesOperacao = operacao ? operacao.pelotoes.map(normalizarPelotaoResenha_) : [];
  const operacaoDupla = pelotoesOperacao.includes('1º PEL') && pelotoesOperacao.includes('2º PEL');
  const pel = normalizarPelotaoResenha_(pelotao);
  const duplicada = operacaoDupla
    ? existentes.some(function(existente) { return existente === pel || !['1º PEL', '2º PEL'].includes(existente); })
    : true;
  if (duplicada) throw new Error('RESENHA JÁ ENVIADA PARA ESSA DATA');
}

function epapiEnviarResenhaEscala_(params) {
  const autenticacao = autenticarUsuarioParaAcao_(params, 'RESENHA_DE_ESCALAS', 'inserir');
  if (!autenticacao.ok) return respostaFalha_(autenticacao.erro);
  const payload = parseJsonParam_(params.payload || params.dados, {});
  const validacaoViaturas = validarViaturasPayloadResenha_(payload);
  if (!validacaoViaturas.ok) return respostaFalha_(validacaoViaturas.erro);
  const p = autenticacao.policial;
  const pelotaoInformado = payload.pelotao || params.pelotao || p.pelotao || '';
  const validacaoEscopo = validarEscopoEnvioResenha_(p, autenticacao.permissao, pelotaoInformado);
  if (!validacaoEscopo.ok) return respostaFalha_(validacaoEscopo.erro);
  const pelotao = validacaoEscopo.pelotao;
  payload.pelotao = pelotao;
  const id = 'RES-' + Utilities.formatDate(new Date(), FT_API.timezone, 'yyyyMMddHHmmss') + '-' + normalizarLoginRE_(p.re) + '-' + Utilities.getUuid().slice(0, 8);
  const dataEscala = payload.dataEscala || params.dataEscala || '';
  const horario = payload.horarioEscala || '';
  const obs = payload.observacoes || '';
  const itens = payload.itens || [];
  const complementos = payload.complementos || [];
  const lock = LockService.getScriptLock();
  const gravados = [];
  try {
    lock.waitLock(30000);
    epVerificarBloqueioResenhaData_(dataEscala, pelotao);
    epValidarFolgasResenha_(payload,p,id,null);
    const abaMain = obterAbaResenhasApp_();
    const linhaMain = Math.max(2, abaMain.getLastRow() + 1);
    abaMain.getRange(linhaMain, 1, 1, 15).setValues([[
      id, new Date(), dataEscala, pelotao, horario, p.re,
      [p.postoGrad, p.nomeGuerra].filter(Boolean).join(' '), 'ENVIADA', obs,
      JSON.stringify(payload), '', '', '', 'PENDENTE', ''
    ]]);
    gravados.push({ aba: abaMain, linha: linhaMain, quantidade: 1, colunas: 15 });

    if (itens.length) {
      const abaItens = obterAbaItensResenhasApp_();
      const linhaItens = Math.max(2, abaItens.getLastRow() + 1);
      const linhasItens = itens.map(function(item) {
        return [id, item.tipo || '', item.grupo || '', item.campo || '', item.valor || ''];
      });
      abaItens.getRange(linhaItens, 1, linhasItens.length, 5).setValues(linhasItens);
      gravados.push({ aba: abaItens, linha: linhaItens, quantidade: linhasItens.length, colunas: 5 });
    }

    if (complementos.length) {
      const abaComp = obterAbaComplementosResenhasApp_();
      const linhaComp = Math.max(2, abaComp.getLastRow() + 1);
      const linhasComp = complementos.map(function(c) { return [id, c.tipo || '', c.texto || '']; });
      abaComp.getRange(linhaComp, 1, linhasComp.length, 3).setValues(linhasComp);
      gravados.push({ aba: abaComp, linha: linhaComp, quantidade: linhasComp.length, colunas: 3 });
    }
  } catch (erro) {
    /* Se uma das tr\u00EAs grava\u00E7\u00F5es falhar, remove os blocos j\u00E1 escritos para n\u00E3o
       deixar novamente uma resenha \u00F3rf\u00E3 ou parcialmente registrada. */
    gravados.reverse().forEach(function(item) {
      try { item.aba.getRange(item.linha, 1, item.quantidade, item.colunas).clearContent(); } catch (ignorado) {}
    });
    if (erro && erro.message === 'RESENHA JÁ ENVIADA PARA ESSA DATA') return respostaFalha_(erro.message);
    return respostaFalha_('N\u00E3o foi poss\u00EDvel gravar a Resenha de Escala de forma \u00EDntegra. Nenhum lan\u00E7amento parcial foi mantido. ' + (erro && erro.message ? erro.message : erro));
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
  registrarHistorico_({ usuario: p, modulo: 'RESENHA_DE_ESCALAS', acao: 'ENVIAR_RESENHA_ESCALA', idRegistro: id, observacao: 'Resenha enviada pelo app.', status: 'OK' });
  return respostaOk_({ mensagem: 'Resenha de escala enviada com sucesso.', id: id });
}



function epapiEditarResenhaEscala_(params) {
  const autenticacao = autenticarUsuarioParaAcao_(params, 'RESENHA_DE_ESCALAS', 'visualizar');
  if (!autenticacao.ok) return respostaFalha_(autenticacao.erro);
  const id = String(params.id || '').trim();
  if (!id) return respostaFalha_('Informe a resenha que será editada.');
  const podeAdministrar = perfilAdminResenhaEscala_(autenticacao.policial);
  if (!podeAdministrar && (!perfilCmtEquipeResenha_(autenticacao.policial) ||
      (!autenticacao.permissao.inserir && !autenticacao.permissao.acessoTotal))) {
    return respostaFalha_('A edição é restrita ao Sgt autor enquanto a resenha estiver aberta, ou ao Gestor/Enc Adm.');
  }

  const payload = parseJsonParam_(params.payload || params.dados, {});
  const validacaoViaturas = validarViaturasPayloadResenha_(payload);
  if (!validacaoViaturas.ok) return respostaFalha_(validacaoViaturas.erro);
  if (!podeAdministrar) {
    const escopo = validarEscopoEnvioResenha_(autenticacao.policial, autenticacao.permissao,
      payload.pelotao || autenticacao.policial.pelotao || '');
    if (!escopo.ok) return respostaFalha_(escopo.erro);
    payload.pelotao = escopo.pelotao;
  } else {
    payload.pelotao = normalizarPelotaoResenha_(payload.pelotao || '');
    if (!['1º PEL','2º PEL'].includes(payload.pelotao)) return respostaFalha_('Pelotão inválido na resenha.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const abaAppExistente = obterAbaSeExistir_(FT_API.abas.resenhasApp);
    if (abaAppExistente) obterAbaResenhasApp_();
    const abaLegadaExistente = obterAbaSeExistir_('RESENHAS_ESCALAS');
    if (abaLegadaExistente) obterOuCriarAbaComCabecalho_('RESENHAS_ESCALAS', cabecalhoResenhasApp_());
    const abasMain = [obterAbaSeExistir_(FT_API.abas.resenhasApp), obterAbaSeExistir_('RESENHAS_ESCALAS')]
      .filter(function(aba, indice, lista) { return aba && lista.indexOf(aba) === indice; });
    let principal = null;
    for (let a = 0; a < abasMain.length && !principal; a++) {
      const aba = abasMain[a];
      if (aba.getLastRow() < 2) continue;
      const ids = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getDisplayValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0] || '').trim() === id) {
          const linha = i + 2;
          const valores = aba.getRange(linha, 1, 1, 19).getValues()[0];
          principal = { aba: aba, linha: linha, valores: valores };
          break;
        }
      }
    }
    if (!principal) return respostaFalha_('Resenha não localizada. Atualize a página e tente novamente.');

    const anterior = principal.valores;
    const statusAnterior = String(anterior[13] || anterior[7] || 'PENDENTE');
    if (!usuarioPodeEditarResenhaEscala_(autenticacao, anterior[5], anterior[3], statusAnterior)) {
      return respostaFalha_('Esta resenha não pode mais ser editada ou pertence a outro pelotão. Resenhas aprovadas ou recusadas ficam bloqueadas.');
    }
    if (normalizarPelotaoResenha_(payload.pelotao) !== normalizarPelotaoResenha_(anterior[3])) {
      return respostaFalha_('A edição deve manter o pelotão da resenha original.');
    }

    epValidarFolgasResenha_(payload,autenticacao.policial,id,anterior);
    const abasItens = [obterAbaSeExistir_(FT_API.abas.resenhasAppItens), obterAbaSeExistir_('RESENHAS_ESCALAS_ITENS')]
      .filter(function(aba, indice, lista) { return aba && lista.indexOf(aba) === indice; });
    const abasComp = [obterAbaSeExistir_(FT_API.abas.resenhasAppComplementos), obterAbaSeExistir_('RESENHAS_ESCALAS_COMPLEMENTOS')]
      .filter(function(aba, indice, lista) { return aba && lista.indexOf(aba) === indice; });
    const itensAnteriores = lerLinhasFilhasResenha_(abasItens, id, 5);
    const complementosAnteriores = lerLinhasFilhasResenha_(abasComp, id, 3);
    const dataEdicao = new Date();
    const nomeEditor = [autenticacao.policial.postoGrad, autenticacao.policial.re, autenticacao.policial.nomeGuerra].filter(Boolean).join(' ');
    const abaHistorico = obterAbaHistoricoEdicoesResenha_();
    const linhaHistorico = abaHistorico.getLastRow() + 1;
    abaHistorico.appendRow([
      Utilities.getUuid(), id, dataEdicao, autenticacao.policial.re || '', nomeEditor,
      String(anterior[7] || ''), String(anterior[13] || ''), String(anterior[14] || ''),
      String(anterior[9] || ''), JSON.stringify(itensAnteriores.map(function(item) { return item.valores; })),
      JSON.stringify(complementosAnteriores.map(function(item) { return item.valores; })), 'EM_ANDAMENTO'
    ]);

    const atualizado = anterior.slice(0, 19);
    while (atualizado.length < 19) atualizado.push('');
    atualizado[2] = payload.dataEscala || anterior[2];
    atualizado[3] = payload.pelotao;
    atualizado[4] = payload.horarioEscala || '';
    atualizado[7] = anterior[7] || 'ENVIADA';
    atualizado[8] = payload.observacoes || '';
    atualizado[9] = JSON.stringify(payload);
    atualizado[13] = anterior[13] || 'PENDENTE';
    atualizado[15] = dataEdicao;
    atualizado[16] = autenticacao.policial.re || '';
    atualizado[17] = nomeEditor;
    atualizado[18] = Number(anterior[18] || 0) + 1;

    try {
      removerLinhasFilhasResenha_(abasItens, id);
      removerLinhasFilhasResenha_(abasComp, id);
      gravarFilhosResenha_(id, payload.itens || [], payload.complementos || []);
      principal.aba.getRange(principal.linha, 1, 1, 19).setValues([atualizado]);
      SpreadsheetApp.flush();
      try { abaHistorico.getRange(linhaHistorico, 12).setValue('CONCLUIDA'); } catch (erroAuditoria) {}
    } catch (erroEdicao) {
      try {
        removerLinhasFilhasResenha_(abasItens, id);
        removerLinhasFilhasResenha_(abasComp, id);
        itensAnteriores.forEach(function(item) {
          item.aba.getRange(item.aba.getLastRow() + 1, 1, 1, 5).setValues([item.valores]);
        });
        complementosAnteriores.forEach(function(item) {
          item.aba.getRange(item.aba.getLastRow() + 1, 1, 1, 3).setValues([item.valores]);
        });
        principal.aba.getRange(principal.linha, 1, 1, 19).setValues([anterior]);
        try { abaHistorico.getRange(linhaHistorico, 12).setValue('REVERTIDA'); } catch (erroAuditoria) {}
      } catch (erroRollback) {}
      throw erroEdicao;
    }

    registrarHistorico_({
      usuario: autenticacao.policial,
      modulo: 'RESENHA_DE_ESCALAS',
      acao: 'EDITAR_RESENHA_ESCALA',
      idRegistro: id,
      valorAnterior: statusAnterior,
      valorNovo: String(atualizado[13] || atualizado[7]),
      observacao: 'Resenha editada no mesmo registro. Versão anterior preservada em RESENHAS_APP_REVISOES; status decisório mantido.',
      status: 'OK'
    });
    return respostaOk_({ id: id, mensagem: 'Resenha atualizada com sucesso.', status: atualizado[13] || atualizado[7], dataHoraEdicao: formatarDataHora_(dataEdicao), editadoPor: nomeEditor, numeroEdicoes: atualizado[18] });
  } catch (erro) {
    return respostaFalha_('Não foi possível salvar a edição da resenha. A versão anterior foi preservada. ' + (erro && erro.message ? erro.message : erro));
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}


apiEnviarResenhaEscala_=epapiEnviarResenhaEscala_;
apiEditarResenhaEscala_=epapiEditarResenhaEscala_;
