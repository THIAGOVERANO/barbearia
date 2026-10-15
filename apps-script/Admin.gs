// ============================================================
// PAINEL DE ADMINISTRAÇÃO (uso só da barbearia)
// A senha NÃO fica neste arquivo: ela fica em "Propriedades do script".
// ============================================================

const STATUS_VALIDOS = ['Confirmado', 'Concluído', 'Cancelado', 'Não compareceu'];
const SENHA_ERRADA = { ok: false, erro: 'Senha incorreta. Tente de novo.' };

// Abre o painel quando alguém acessa o link com ?page=admin
function paginaAdmin() {
  return HtmlService.createTemplateFromFile('admin')
    .evaluate()
    .setTitle('Agenda Zure · Painel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function senhaCorreta(senha) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_SENHA');
  return !!real && String(senha || '') === real;
}

function entrarNoPainel(senha) {
  return senhaCorreta(senha) ? { ok: true } : SENHA_ERRADA;
}

function dataValida(d) {
  return /^\d{4}-\d{2}-\d{2}$/.test(d || '');
}

function horaValida(h) {
  return /^\d{2}:\d{2}$/.test(h || '');
}

function somarDias(dataStr, n) {
  const p = dataStr.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)).toISOString().slice(0, 10);
}

function linhaExiste(aba, linha) {
  return !!aba && Number.isInteger(linha) && linha >= 2 && linha <= aba.getLastRow();
}

// Lê uma aba guardando o número da linha (para poder mudar ou apagar depois)
function linhasDaAba(nomeAba) {
  const aba = SpreadsheetApp.getActive().getSheetByName(nomeAba);
  if (!aba || aba.getLastRow() < 2) return [];
  const valores = aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getValues();
  return valores.map((dados, i) => ({ linha: i + 2, dados: dados }));
}

function agendamentosDoDia(data) {
  return linhasDaAba(ABA_AGENDAMENTOS)
    .filter(x => x.dados[1] && textoData(x.dados[1]) === data)
    .map(x => ({
      linha: x.linha,
      inicio: textoHora(x.dados[2]),
      fim: textoHora(x.dados[3]),
      servico: String(x.dados[4]),
      barbeiro: String(x.dados[5]),
      cliente: String(x.dados[6]),
      telefone: String(x.dados[7]),
      status: String(x.dados[8])
    }))
    .sort((a, b) => minutos(a.inicio) - minutos(b.inicio));
}

function bloqueiosDoDia(data) {
  return linhasDaAba(ABA_BLOQUEIOS)
    .filter(x => x.dados[0] && textoData(x.dados[0]) === data)
    .map(x => ({
      linha: x.linha,
      inicio: textoHora(x.dados[1]),
      fim: textoHora(x.dados[2]),
      barbeiro: String(x.dados[3] || TODOS),
      motivo: String(x.dados[4] || '')
    }))
    .sort((a, b) => minutos(a.inicio) - minutos(b.inicio));
}

// Tudo o que o painel mostra de um dia
function adminDia(senha, data) {
  if (!senhaCorreta(senha)) return SENHA_ERRADA;
  if (!dataValida(data)) return { ok: false, erro: 'Data inválida.' };
  return {
    ok: true,
    data: data,
    agendamentos: agendamentosDoDia(data),
    bloqueios: bloqueiosDoDia(data),
    barbeiros: CONFIG.barbeiros,
    servicos: SERVICOS.map(s => ({ id: s.id, nome: s.nome }))
  };
}

// Agendamento feito pelo painel (encaixe por telefone, por exemplo).
// Aceita qualquer horário, desde que não bata com outro agendamento ou bloqueio.
function adminAgendar(senha, data, hora, servicoId, barbeiro, nome, telefone) {
  if (!senhaCorreta(senha)) return SENHA_ERRADA;
  const servico = SERVICOS.find(s => s.id === servicoId);
  if (!servico) return { ok: false, erro: 'Serviço inválido.' };
  if (!dataValida(data) || !horaValida(hora)) return { ok: false, erro: 'Data ou horário inválidos.' };
  const escolha = barbeiro || TANTO_FAZ;
  if (!escolhaValida(escolha)) return { ok: false, erro: 'Barbeiro inválido.' };
  const nomeLimpo = String(nome || '').trim();
  const fone = String(telefone || '').replace(/\D/g, '');
  if (nomeLimpo.length < 2) return { ok: false, erro: 'Informe o nome do cliente.' };
  if (fone.length < 10 || fone.length > 11) return { ok: false, erro: 'Informe o telefone com DDD.' };

  const trava = LockService.getScriptLock();
  trava.waitLock(15000);
  try {
    const ini = minutos(hora);
    const fim = ini + servico.duracao;
    const ocupados = intervalosOcupados(data);
    const barbs = barbeirosPara(servico, escolha);
    const livre = barbs.find(b => barbeiroLivre(ocupados, b, ini, fim));
    if (!livre) return { ok: false, erro: 'Esse horário já está ocupado para o barbeiro escolhido.' };

    const aba = SpreadsheetApp.getActive().getSheetByName(ABA_AGENDAMENTOS);
    aba.appendRow([
      Utilities.getUuid().slice(0, 8), data, horaDeMinutos(ini), horaDeMinutos(fim),
      servico.nome, livre, nomeLimpo, fone, 'Confirmado',
      Utilities.formatDate(new Date(), CONFIG.fuso, 'dd/MM/yyyy HH:mm') + ' (painel)'
    ]);
    ordenarAba(ABA_AGENDAMENTOS, 1, 2);
    return { ok: true, barbeiro: livre };
  } finally {
    trava.releaseLock();
  }
}

// Resumo de 7 dias a partir de uma data (para a faixa da semana)
function adminSemana(senha, dataInicio) {
  if (!senhaCorreta(senha)) return SENHA_ERRADA;
  if (!dataValida(dataInicio)) return { ok: false, erro: 'Data inválida.' };
  const dias = [];
  for (let i = 0; i < 7; i++) {
    const dia = somarDias(dataInicio, i);
    dias.push({
      data: dia,
      agendamentos: agendamentosDoDia(dia).filter(a => a.status !== 'Cancelado').length,
      bloqueios: bloqueiosDoDia(dia).length
    });
  }
  return { ok: true, dias: dias };
}

// Muda o status de um agendamento (Confirmado, Concluído, Cancelado, Não compareceu)
function adminMudarStatus(senha, linha, status) {
  if (!senhaCorreta(senha)) return SENHA_ERRADA;
  if (STATUS_VALIDOS.indexOf(status) === -1) return { ok: false, erro: 'Status inválido.' };
  const aba = SpreadsheetApp.getActive().getSheetByName(ABA_AGENDAMENTOS);
  if (!linhaExiste(aba, linha)) return { ok: false, erro: 'Agendamento não encontrado.' };
  aba.getRange(linha, 9).setValue(status);
  return { ok: true };
}

// Bloqueia um período (almoço, folga, compromisso)
function adminBloquear(senha, data, inicio, fim, barbeiro, motivo) {
  if (!senhaCorreta(senha)) return SENHA_ERRADA;
  if (!dataValida(data) || !horaValida(inicio) || !horaValida(fim)) {
    return { ok: false, erro: 'Data ou horário inválidos.' };
  }
  if (minutos(fim) <= minutos(inicio)) return { ok: false, erro: 'O horário final precisa ser depois do inicial.' };
  if (barbeiro !== TODOS && CONFIG.barbeiros.indexOf(barbeiro) === -1) return { ok: false, erro: 'Barbeiro inválido.' };

  const trava = LockService.getScriptLock();
  trava.waitLock(15000);
  try {
    const aba = SpreadsheetApp.getActive().getSheetByName(ABA_BLOQUEIOS);
    aba.appendRow([data, inicio, fim, barbeiro, String(motivo || '').slice(0, 80)]);
    ordenarAba(ABA_BLOQUEIOS, 0, 1);
  } finally {
    trava.releaseLock();
  }

  // Avisa se o bloqueio cai em cima de agendamentos que já existem
  const ini = minutos(inicio);
  const fi = minutos(fim);
  const conflitos = agendamentosDoDia(data).filter(a =>
    a.status !== 'Cancelado' &&
    minutos(a.inicio) < fi && ini < minutos(a.fim) &&
    (barbeiro === TODOS || a.barbeiro === barbeiro)
  ).length;
  return { ok: true, conflitos: conflitos };
}

// Remove um bloqueio (pelo número da linha que o painel mostra)
function adminRemoverBloqueio(senha, linha) {
  if (!senhaCorreta(senha)) return SENHA_ERRADA;
  const aba = SpreadsheetApp.getActive().getSheetByName(ABA_BLOQUEIOS);
  if (!linhaExiste(aba, linha)) return { ok: false, erro: 'Bloqueio não encontrado.' };
  aba.deleteRow(linha);
  return { ok: true };
}
