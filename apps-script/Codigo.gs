// ============================================================
// AGENDA ZURE - código do servidor (Google Apps Script)
// Só mexa na parte de cima (CONFIG e SERVICOS) para mudar horários,
// barbeiros e preços. O resto deixe como está.
// ============================================================

const CONFIG = {
  fuso: 'America/Sao_Paulo',
  abertura: '08:00',        // primeiro horário que o cliente pode escolher
  fechamento: '18:00',      // horário de fechamento
  intervaloMinutos: 30,     // os horários aparecem de 30 em 30 minutos
  diasFechados: [0],        // 0 = domingo. Para fechar também na segunda, use [0, 1]
  diasParaFrente: 30,       // quantos dias à frente o cliente pode marcar
  barbeiros: ['Gustavo', 'Natan'],
  whatsapp: '5562994964571',
  nomeBarbearia: 'Zure Barbearia'
};

// duracao = minutos. preco = reais. barbeiro = só se o serviço for com um barbeiro fixo.
const SERVICOS = [
  { id: 'corte-militar', nome: 'Corte militar', duracao: 30, preco: 40, grupo: 'Barbearia' },
  { id: 'barba', nome: 'Barba', duracao: 30, preco: 50, grupo: 'Barbearia' },
  { id: 'corte', nome: 'Corte de cabelo', duracao: 60, preco: 60, grupo: 'Barbearia' },
  { id: 'corte-natan', nome: 'Corte com Natan', duracao: 60, preco: 45, grupo: 'Barbearia', barbeiro: 'Natan' },
  { id: 'corte-gustavo', nome: 'Corte com Gustavo', duracao: 60, preco: 65, grupo: 'Barbearia', barbeiro: 'Gustavo' },
  { id: 'combo', nome: 'Combo cabelo e barba', duracao: 90, preco: 100, grupo: 'Barbearia' },
  { id: 'sobrancelha', nome: 'Sobrancelha na navalha', duracao: 10, preco: 15, grupo: 'Barbearia' },
  { id: 'pe-cabelo', nome: 'Pé do cabelo', duracao: 15, preco: 20, grupo: 'Barbearia' },
  { id: 'corte-selagem', nome: 'Corte e selagem', duracao: 95, preco: 140, grupo: 'Selagem e tratamentos' },
  { id: 'corte-barba-selagem', nome: 'Corte, barba e selagem', duracao: 100, preco: 180, grupo: 'Selagem e tratamentos' },
  { id: 'selagem', nome: 'Selagem ou progressiva', duracao: 60, preco: 90, grupo: 'Selagem e tratamentos' },
  { id: 'selagem-hidratacao', nome: 'Selagem e hidratação express', duracao: 85, preco: 115, grupo: 'Selagem e tratamentos' },
  { id: 'hidratacao-express', nome: 'Hidratação express', duracao: 20, preco: 40, grupo: 'Selagem e tratamentos', descricao: 'creme de manteiga, argan ou oliva com ozônio' },
  { id: 'hidratacao-shock', nome: 'Hidratação shock', duracao: 40, preco: 60, grupo: 'Selagem e tratamentos', descricao: 'maciez, menos frizz e mais volume' }
];

const ABA_AGENDAMENTOS = 'Agendamentos';
const ABA_BLOQUEIOS = 'Bloqueios';
const TODOS = 'Todos';
const TANTO_FAZ = 'Tanto faz';

// ------------------------------------------------------------
// Função de preparação: rode UMA vez (veja o passo a passo)
// ------------------------------------------------------------
function preparar() {
  const ss = SpreadsheetApp.getActive();

  let ag = ss.getSheetByName(ABA_AGENDAMENTOS) || ss.insertSheet(ABA_AGENDAMENTOS);
  if (ag.getLastRow() === 0) {
    ag.appendRow(['ID', 'Data', 'Início', 'Fim', 'Serviço', 'Barbeiro', 'Cliente', 'Telefone', 'Status', 'Feito em']);
    ag.getRange('1:1').setFontWeight('bold');
  }
  ag.getRange('B:D').setNumberFormat('@');

  let bl = ss.getSheetByName(ABA_BLOQUEIOS) || ss.insertSheet(ABA_BLOQUEIOS);
  if (bl.getLastRow() === 0) {
    bl.appendRow(['Data', 'Início', 'Fim', 'Barbeiro', 'Motivo']);
    bl.getRange('1:1').setFontWeight('bold');
  }
  bl.getRange('A:C').setNumberFormat('@');
}

// ------------------------------------------------------------
// Entrada do site (pedidos de leitura)
// ------------------------------------------------------------
function doGet(e) {
  if (e.parameter.page === 'admin') return paginaAdmin();
  const acao = (e.parameter.action || 'config');
  try {
    if (acao === 'config') {
      return resposta({
        ok: true,
        servicos: SERVICOS,
        barbeiros: CONFIG.barbeiros,
        abertura: CONFIG.abertura,
        fechamento: CONFIG.fechamento,
        intervaloMinutos: CONFIG.intervaloMinutos,
        diasFechados: CONFIG.diasFechados,
        diasParaFrente: CONFIG.diasParaFrente,
        whatsapp: CONFIG.whatsapp,
        nomeBarbearia: CONFIG.nomeBarbearia
      });
    }
    if (acao === 'horarios') {
      const horarios = horariosLivres(e.parameter.data, e.parameter.servico, e.parameter.barbeiro || TANTO_FAZ);
      return resposta({ ok: true, horarios: horarios });
    }
    return resposta({ ok: false, erro: 'Ação desconhecida.' });
  } catch (err) {
    return resposta({ ok: false, erro: String(err.message || err) });
  }
}

// ------------------------------------------------------------
// Entrada do site (pedido de agendamento)
// ------------------------------------------------------------
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    return resposta(criarAgendamento(dados));
  } catch (err) {
    return resposta({ ok: false, erro: 'Não foi possível confirmar. Tente de novo ou chame a barbearia no WhatsApp.' });
  }
}

// ------------------------------------------------------------
// Regras de horário
// ------------------------------------------------------------
function horariosLivres(dataStr, servicoId, escolha) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dataStr || '')) throw new Error('Data inválida.');
  const servico = SERVICOS.find(s => s.id === servicoId);
  if (!servico) throw new Error('Serviço inválido.');
  if (!escolhaValida(escolha)) throw new Error('Barbeiro inválido.');

  const hoje = Utilities.formatDate(new Date(), CONFIG.fuso, 'yyyy-MM-dd');
  if (dataStr < hoje) return [];

  const partes = dataStr.split('-').map(Number);
  const diaSemana = new Date(Date.UTC(partes[0], partes[1] - 1, partes[2])).getUTCDay();
  if (CONFIG.diasFechados.indexOf(diaSemana) !== -1) return [];

  const agora = new Date();
  const minAgora = minutos(Utilities.formatDate(agora, CONFIG.fuso, 'HH:mm'));
  const ocupadosDia = intervalosOcupados(dataStr);
  const abre = minutos(CONFIG.abertura);
  const fecha = minutos(CONFIG.fechamento);
  const barbs = barbeirosPara(servico, escolha);

  const livres = [];
  for (let ini = abre; ini + servico.duracao <= fecha; ini += CONFIG.intervaloMinutos) {
    if (dataStr === hoje && ini <= minAgora) continue;
    const fim = ini + servico.duracao;
    const alguemLivre = barbs.some(b => barbeiroLivre(ocupadosDia, b, ini, fim));
    if (alguemLivre) livres.push(horaDeMinutos(ini));
  }
  return livres;
}

function criarAgendamento(p) {
  const servico = SERVICOS.find(s => s.id === p.servico);
  if (!servico) return { ok: false, erro: 'Serviço inválido.' };

  const nome = String(p.nome || '').trim();
  const telefone = String(p.telefone || '').replace(/\D/g, '');
  if (nome.length < 2) return { ok: false, erro: 'Informe seu nome.' };
  if (telefone.length < 10 || telefone.length > 11) return { ok: false, erro: 'Informe o telefone com DDD.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(p.data || '') || !/^\d{2}:\d{2}$/.test(p.hora || '')) {
    return { ok: false, erro: 'Data ou horário inválidos.' };
  }
  const escolha = p.barbeiro || TANTO_FAZ;
  if (!escolhaValida(escolha)) return { ok: false, erro: 'Barbeiro inválido.' };

  // Trava para duas pessoas não marcarem o mesmo horário ao mesmo tempo
  const trava = LockService.getScriptLock();
  trava.waitLock(15000);
  try {
    const livres = horariosLivres(p.data, servico.id, escolha);
    if (livres.indexOf(p.hora) === -1) {
      return { ok: false, erro: 'Esse horário acabou de ser ocupado. Escolha outro, por favor.' };
    }

    const ini = minutos(p.hora);
    const fim = ini + servico.duracao;
    const ocupadosDia = intervalosOcupados(p.data);
    const barbs = barbeirosPara(servico, escolha);
    const barbeiro = barbs.find(b => barbeiroLivre(ocupadosDia, b, ini, fim));

    const aba = SpreadsheetApp.getActive().getSheetByName(ABA_AGENDAMENTOS);
    const id = Utilities.getUuid().slice(0, 8);
    aba.appendRow([
      id, p.data, horaDeMinutos(ini), horaDeMinutos(fim),
      servico.nome, barbeiro, nome, telefone, 'Confirmado',
      Utilities.formatDate(new Date(), CONFIG.fuso, 'dd/MM/yyyy HH:mm')
    ]);
    ordenarAba(ABA_AGENDAMENTOS, 1, 2);

    return { ok: true, id: id, servico: servico.nome, barbeiro: barbeiro, data: p.data, hora: p.hora };
  } finally {
    trava.releaseLock();
  }
}

// ------------------------------------------------------------
// Ajudantes
// ------------------------------------------------------------
function escolhaValida(escolha) {
  return escolha === TANTO_FAZ || CONFIG.barbeiros.indexOf(escolha) !== -1;
}

function barbeirosPara(servico, escolha) {
  if (servico.barbeiro) return [servico.barbeiro];
  if (escolha && escolha !== TANTO_FAZ) return [escolha];
  return CONFIG.barbeiros;
}

// Junta agendamentos (que não estão cancelados) e bloqueios de um dia
function intervalosOcupados(dataStr) {
  const lista = [];

  lerLinhas(ABA_AGENDAMENTOS).forEach(l => {
    // colunas: 0 ID, 1 Data, 2 Início, 3 Fim, 4 Serviço, 5 Barbeiro, 6 Cliente, 7 Telefone, 8 Status
    if (!l[1] || textoData(l[1]) !== dataStr) return;
    if (String(l[8]).trim() === 'Cancelado') return;
    lista.push({
      barbeiro: String(l[5]).trim(),
      ini: minutos(textoHora(l[2])),
      fim: minutos(textoHora(l[3]))
    });
  });

  lerLinhas(ABA_BLOQUEIOS).forEach(l => {
    // colunas: 0 Data, 1 Início, 2 Fim, 3 Barbeiro (ou Todos), 4 Motivo
    if (!l[0] || textoData(l[0]) !== dataStr) return;
    lista.push({
      barbeiro: String(l[3]).trim() || TODOS,
      ini: minutos(textoHora(l[1])),
      fim: minutos(textoHora(l[2]))
    });
  });

  return lista;
}

function barbeiroLivre(ocupadosDia, barbeiro, ini, fim) {
  return !ocupadosDia.some(o =>
    (o.barbeiro === barbeiro || o.barbeiro === TODOS) && o.ini < fim && ini < o.fim
  );
}

function lerLinhas(nomeAba) {
  const aba = SpreadsheetApp.getActive().getSheetByName(nomeAba);
  if (!aba) return [];
  return aba.getDataRange().getValues().slice(1);
}

// O Google Sheets às vezes transforma data e hora em "data de verdade".
// Estas funções sempre devolvem texto no formato certo.
function textoData(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.fuso, 'yyyy-MM-dd');
  return String(v).trim();
}

function textoHora(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.fuso, 'HH:mm');
  return String(v).trim();
}

function minutos(hora) {
  const partes = String(hora).split(':').map(Number);
  return partes[0] * 60 + partes[1];
}

function horaDeMinutos(total) {
  const h = Math.floor(total / 60);
  const m = total % 60;
  return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
}

// Coloca as linhas em ordem de data e horário (linha 1 é o cabeçalho e fica no lugar).
// indiceData e indiceHora são as posições da coluna (0 = primeira coluna).
// Linhas vazias vão para o final.
function ordenarAba(nomeAba, indiceData, indiceHora) {
  const aba = SpreadsheetApp.getActive().getSheetByName(nomeAba);
  if (!aba) return;
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return;
  const area = aba.getRange(2, 1, ultimaLinha - 1, aba.getLastColumn());
  const linhas = area.getValues();
  linhas.sort((a, b) => {
    if (!a[indiceData] && !b[indiceData]) return 0;
    if (!a[indiceData]) return 1;
    if (!b[indiceData]) return -1;
    const dataA = textoData(a[indiceData]), dataB = textoData(b[indiceData]);
    if (dataA !== dataB) return dataA < dataB ? -1 : 1;
    return minutos(textoHora(a[indiceHora])) - minutos(textoHora(b[indiceHora]));
  });
  area.setValues(linhas);
}

function ordenarTudo() {
  ordenarAba(ABA_AGENDAMENTOS, 1, 2);
  ordenarAba(ABA_BLOQUEIOS, 0, 1);
}

// Cria o menu "Agenda" na planilha, para ordenar na hora quando você digitar uma linha à mão
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Agenda').addItem('Ordenar por data', 'ordenarTudo').addToUi();
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
