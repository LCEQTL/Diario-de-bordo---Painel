/* =====================================================================
   DIÁRIO DE BORDO · painel do gestor e do admin
   Páginas: #/dashboard  #/registros  #/pessoas  #/estrutura
   ===================================================================== */
(() => {
  "use strict";
  const CFG = window.PAINEL_CONFIG || {};
  const sb = window.supabase.createClient(CFG.SUPABASE_URL, CFG.SUPABASE_ANON_KEY);
  const TZ = "America/Sao_Paulo";
  const TXT = { feita: "Feita", sem_demanda: "Sem demanda", nao_feita: "Não feita" };
  const PILL = { feita: "p-ok", sem_demanda: "p-sem", nao_feita: "p-nao" };
  const PER = ["Diária", "Semanal", "Mensal", "Eventual"];

  const $ = (s, r = document) => r.querySelector(s);
  const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const norm = (v) => String(v ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").trim().toLowerCase();
  const pct = (v, d = 0) => (v === null || v === undefined ? "—" : (Number(v) * 100).toFixed(d).replace(".", ",") + "%");
  const num = (v, d = 1) => (v === null || v === undefined ? "—" : Number(v).toFixed(d).replace(".", ","));
  const mb = (b) => (b / 1048576).toFixed(b > 104857600 ? 0 : 1).replace(".", ",") + " MB";
  const hojeISO = () => new Intl.DateTimeFormat("en-CA", { timeZone: TZ }).format(new Date());
  const addDias = (iso, n) => { const d = new Date(iso + "T12:00:00Z"); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const br = (iso) => iso ? iso.slice(8, 10) + "/" + iso.slice(5, 7) + "/" + iso.slice(0, 4) : "";
  const dm = (iso) => iso.slice(8, 10) + "/" + iso.slice(5, 7);
  const horaDe = (ts) => new Intl.DateTimeFormat("pt-BR", { timeZone: TZ, hour: "2-digit", minute: "2-digit" }).format(new Date(ts));
  const ls = { get(k, d) { try { const v = localStorage.getItem(k); return v === null ? d : JSON.parse(v); } catch { return d; } },
               set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* ok */ } } };

  const S = { user: null, eu: null, setores: [], cargos: [], rota: "dashboard", msg: "",
    dash: ls.get("painel_dash", { periodo: "mes", ini: "", fim: "", setor: "", cargo: "", tarefa: "", pessoa: "", aba: "pessoa" }),
    reg: { ini: addDias(hojeISO(), -6), fim: hojeISO(), setor: "", pessoa: "", status: "", q: "", ev: "", pag: 0 },
    pes: { q: "", setor: "", perfil: "", ativo: "1" },
    est: { setor: null, cargo: null } };

  const isAdmin = () => S.eu?.perfil === "admin";
  const setorNome = (id) => S.setores.find((s) => s.id === id)?.nome || "—";
  const setorCor = (id) => S.setores.find((s) => s.id === id)?.cor || "#5F6B7A";
  const cargoNome = (id) => S.cargos.find((c) => c.id === id)?.nome || "—";
  const setorPill = (id) => id ? `<span class="setor"><i style="background:${esc(setorCor(id))}"></i>${esc(setorNome(id))}</span>` : "—";
  const setoresVisiveis = () => isAdmin() ? S.setores : S.setores.filter((s) => s.id === S.eu.setor_id);

  /* ------------------------------------------------------------ avisos e modais */
  let tT;
  function toast(msg, erro = false) {
    const el = $("#toast");
    el.className = "toast" + (erro ? " erro-t" : "");
    el.innerHTML = `<span>${esc(msg)}</span><button type="button" aria-label="Fechar">×</button>`;
    el.hidden = false; el.querySelector("button").onclick = () => (el.hidden = true);
    clearTimeout(tT); tT = setTimeout(() => (el.hidden = true), erro ? 9000 : 4000);
  }
  function modal(html, larga = false) {
    const el = $("#modal");
    el.innerHTML = `<div class="modal${larga ? " larga" : ""}" role="dialog" aria-modal="true">${html}</div>`;
    el.hidden = false;
    const f = el.querySelector("input,select,button.prim"); if (f) f.focus();
  }
  function fechar() { $("#modal").hidden = true; $("#modal").innerHTML = ""; }
  $("#modal").addEventListener("click", (e) => { if (e.target.id === "modal" || e.target.closest("[data-fechar]")) fechar(); });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") fechar(); });

  async function funcao(acao, dados) {
    const { data, error } = await sb.functions.invoke("admin-usuarios", { body: { acao, dados } });
    if (error) {
      let m = error.message;
      try { const j = await error.context.json(); if (j?.erro) m = j.erro; } catch { /* sem corpo */ }
      throw new Error(m);
    }
    if (data?.erro) throw new Error(data.erro);
    return data;
  }
  function baixarXlsx(nome, abas) {
    const wb = XLSX.utils.book_new();
    for (const [titulo, linhas] of abas) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(linhas), titulo.slice(0, 31));
    XLSX.writeFile(wb, nome);
  }
  function lerPlanilha(file) {
    return new Promise((ok, falha) => {
      const r = new FileReader();
      r.onload = () => {
        try {
          const wb = XLSX.read(new Uint8Array(r.result), { type: "array" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const linhas = XLSX.utils.sheet_to_json(ws, { defval: "" }).map((o) => {
            const n = {}; for (const k of Object.keys(o)) n[norm(k).replace(/[^a-z]/g, "")] = String(o[k]).trim(); return n;
          });
          ok(linhas);
        } catch (e) { falha(e); }
      };
      r.onerror = () => falha(r.error);
      r.readAsArrayBuffer(file);
    });
  }

  /* ------------------------------------------------------------ casca */
  function casca(conteudo) {
    const item = (r, t) => `<a href="#/${r}" class="${S.rota === r ? "on" : ""}">${t}</a>`;
    return `<div class="casca"><nav class="menu" aria-label="Menu do painel">
      <div class="marca"><i></i>DIÁRIO DE BORDO</div>
      ${item("dashboard", "Dashboard")}${item("registros", "Registros e evidências")}${item("pessoas", "Pessoas")}${item("estrutura", "Setores, cargos e tarefas")}
      <div class="quem"><b>${esc(S.eu.nome)}</b><small>${isAdmin() ? "Admin · todos os setores" : "Gestor · " + esc(setorNome(S.eu.setor_id))}</small>
        <button type="button" data-act="sair">Sair</button></div>
    </nav><main class="conteudo" id="conteudo">${conteudo}</main></div>`;
  }
  function telaLogin() {
    return `<div class="login"><form class="caixa cartao" id="f-login">
      <div class="marca"><i></i>DIÁRIO DE BORDO · PAINEL</div>
      <h1>Entrar</h1>
      <label class="campo">Usuário<span class="sufixo"><input name="u" autocomplete="username" required autofocus><span>@${esc(CFG.EMAIL_DOMINIO)}</span></span></label>
      <label class="campo">Senha<input name="s" type="password" autocomplete="current-password" required style="height:42px;border:1px solid #C9D2DE;border-radius:10px;padding:0 10px"></label>
      <div class="erro">${esc(S.msg)}</div>
      <button class="btn prim" type="submit" style="height:44px">ENTRAR</button>
      <p>Acesso para gestores e administradores.</p></form></div>`;
  }
  function telaSenha() {
    return `<div class="login"><form class="caixa cartao" id="f-senha">
      <div class="marca"><i></i>DIÁRIO DE BORDO · PAINEL</div>
      <h1>Crie sua senha</h1><p>Você entrou com uma senha temporária. Defina uma só sua para continuar.</p>
      <label class="campo">Nova senha<input name="a" type="password" minlength="8" required autocomplete="new-password"></label>
      <label class="campo">Repita a nova senha<input name="b" type="password" minlength="8" required autocomplete="new-password"></label>
      <div class="erro">${esc(S.msg)}</div>
      <button class="btn prim" type="submit" style="height:44px">Salvar e entrar</button></form></div>`;
  }
  function telaAviso(t, x, link) {
    return `<div class="login"><div class="caixa cartao"><div class="marca"><i></i>DIÁRIO DE BORDO · PAINEL</div><h1>${esc(t)}</h1><p>${esc(x)}</p>
      ${link ? `<a class="btn prim" style="display:flex;align-items:center;justify-content:center;text-decoration:none;height:44px" href="${esc(link)}">Abrir o diário</a>` : ""}
      <button class="btn" type="button" data-act="sair">Sair</button></div></div>`;
  }

  async function iniciarSessao(user) {
    S.user = user;
    const { data: eu } = await sb.from("profiles").select("*").eq("id", user.id).single();
    if (!eu) { $("#app").innerHTML = telaAviso("Sem cadastro", "Seu usuário ainda não tem cadastro no Diário de Bordo."); return; }
    S.eu = eu;
    if (eu.perfil === "colaborador") { $("#app").innerHTML = telaAviso("Este acesso é do diário", "Colaboradores registram as tarefas pelo diário de bordo.", CFG.DIARIO_URL); return; }
    if (eu.must_change_password) { $("#app").innerHTML = telaSenha(); return; }
    await carregarBase();
    rotear();
  }
  async function carregarBase() {
    const [s, c] = await Promise.all([sb.from("setores").select("*").order("nome"), sb.from("cargos").select("*").order("nome")]);
    S.setores = s.data || []; S.cargos = c.data || [];
    if (!isAdmin()) { S.dash.setor = String(S.eu.setor_id); S.reg.setor = String(S.eu.setor_id); S.pes.setor = String(S.eu.setor_id); }
  }
  function rotear() {
    if (!S.eu || S.eu.perfil === "colaborador" || S.eu.must_change_password) return;
    const r = (location.hash.replace("#/", "") || "dashboard").split("?")[0];
    S.rota = ["dashboard", "registros", "pessoas", "estrutura"].includes(r) ? r : "dashboard";
    $("#app").innerHTML = casca('<div class="carregando">Carregando…</div>');
    ({ dashboard: pDashboard, registros: pRegistros, pessoas: pPessoas, estrutura: pEstrutura })[S.rota]();
  }
  window.addEventListener("hashchange", rotear);
  const C = () => $("#conteudo");

  /* ================================================================== DASHBOARD */
  function periodoDatas() {
    const d = S.dash, h = hojeISO();
    if (d.periodo === "hoje") return [h, h];
    if (d.periodo === "7d") return [addDias(h, -6), h];
    if (d.periodo === "mesant") { const p = new Date(h.slice(0, 7) + "-01T12:00:00Z"); p.setUTCMonth(p.getUTCMonth() - 1); const ini = p.toISOString().slice(0, 10); return [ini, addDias(h.slice(0, 7) + "-01", -1)]; }
    if (d.periodo === "custom" && d.ini && d.fim) return [d.ini, d.fim];
    return [h.slice(0, 8) + "01", h];
  }
  let cacheTarefas = {}, cachePessoas = {};
  async function tarefasDoSetor(setor) {
    const k = setor || "todos";
    if (!cacheTarefas[k]) { let q = sb.from("tarefas").select("id,codigo,descricao,setor_id").order("codigo"); if (setor) q = q.eq("setor_id", setor); cacheTarefas[k] = (await q).data || []; }
    return cacheTarefas[k];
  }
  async function pessoasDoSetor(setor) {
    const k = setor || "todos";
    if (!cachePessoas[k]) { let q = sb.from("profiles").select("id,nome,setor_id,cargo_id,perfil").eq("perfil", "colaborador").order("nome"); if (setor) q = q.eq("setor_id", setor); cachePessoas[k] = (await q).data || []; }
    return cachePessoas[k];
  }

  async function pDashboard() {
    const d = S.dash; ls.set("painel_dash", d);
    const [ini, fim] = periodoDatas();
    const setor = d.setor ? Number(d.setor) : null;
    const [tarefas, pessoas] = await Promise.all([tarefasDoSetor(setor), pessoasDoSetor(setor)]);
    const cargos = S.cargos.filter((c) => !setor || c.setor_id === setor);
    const args = { p_inicio: ini, p_fim: fim, p_setor: setor, p_cargo: d.cargo ? Number(d.cargo) : null, p_tarefa: d.tarefa ? Number(d.tarefa) : null, p_user: d.pessoa || null };
    const [res, uso] = await Promise.all([sb.rpc("dashboard", args), isAdmin() ? sb.rpc("uso_armazenamento") : Promise.resolve({ data: null })]);
    if (res.error) { C().innerHTML = `<div class="cartao vazio">Erro ao carregar o dashboard: ${esc(res.error.message)}</div>`; return; }
    const D = res.data, K = D.kpis;
    const periodos = [["hoje", "Hoje"], ["7d", "7 dias"], ["mes", "Mês atual"], ["mesant", "Mês anterior"], ["custom", "Personalizado"]];
    const opt = (v, t, sel) => `<option value="${esc(v)}"${String(sel) === String(v) ? " selected" : ""}>${esc(t)}</option>`;
    const totStatus = K.total || 0;
    const dist = totStatus ? `${pct(K.feitas / totStatus)} · ${pct(K.sem_demanda / totStatus)} · ${pct(K.nao_feitas / totStatus)}` : "—";
    const maxDia = Math.max(1, ...D.por_dia.map((x) => x.feitas + x.sem + x.nao));
    const h = (v) => Math.round(v / maxDia * 190);
    const maxMot = Math.max(1, ...D.motivos.map((m) => m.n));
    const alerta = (r) => { const a = []; if (r.adesao !== null && r.adesao < .8) a.push("adesão baixa"); if (r.pct_nao > .1) a.push("muitas não feitas"); if (r.pct_ev !== null && r.pct_ev < .3 && r.total) a.push("evidência baixa"); return a.length ? `<span class="pill p-nao">${a.join(" · ")}</span>` : `<span class="pill p-ok">ok</span>`; };
    const barra = (v) => `<span class="mini-barra"><i style="width:${Math.round((v || 0) * 100)}%;background:${(v || 0) < .8 ? "var(--nao-bar)" : "var(--azul)"}"></i></span>${pct(v)}`;
    const setorTag = (nome, cor) => nome ? `<span class="setor"><i style="background:${esc(cor)}"></i>${esc(nome)}</span>` : "—";
    const abas = {
      pessoa: { cab: ["Colaborador", "Setor", "Cargo", "Adesão", "Feitas/dia", "Tempo médio", "1º registro", "% não feita", "Com evidência", "Alerta"],
        lin: D.por_pessoa.map((r) => [esc(r.nome), setorTag(r.setor, r.cor), esc(r.cargo || "—"), barra(r.adesao), num(r.feitas_dia), r.tempo != null ? r.tempo + " min" : "—", r.primeiro || "—", pct(r.pct_nao), pct(r.pct_ev), alerta(r)]),
        xls: D.por_pessoa.map((r) => ({ Colaborador: r.nome, Setor: r.setor, Cargo: r.cargo, "Adesão": r.adesao, "Feitas/dia": r.feitas_dia, "Tempo médio (min)": r.tempo, "1º registro": r.primeiro, "% não feita": r.pct_nao, "% com evidência": r.pct_ev, Registros: r.total })) },
      cargo: { cab: ["Cargo", "Setor", "Pessoas", "Adesão média", "Feitas/pessoa/dia", "Tempo médio", "% não feita"],
        lin: D.por_cargo.map((r) => [esc(r.cargo), setorTag(r.setor, r.cor), r.pessoas, barra(r.adesao), num(r.feitas_dia), r.tempo != null ? r.tempo + " min" : "—", pct(r.pct_nao)]),
        xls: D.por_cargo.map((r) => ({ Cargo: r.cargo, Setor: r.setor, Pessoas: r.pessoas, "Adesão média": r.adesao, "Feitas/pessoa/dia": r.feitas_dia, "Tempo médio (min)": r.tempo, "% não feita": r.pct_nao })) },
      tarefa: { cab: ["Código", "Tarefa", "Setor", "Registros", "% feita", "% sem demanda", "% não feita", "Motivo principal", "Tempo médio"],
        lin: D.por_tarefa.map((r) => [`<span class="cod">${esc(r.codigo)}</span>`, esc(r.tarefa), setorTag(r.setor, r.cor), r.registros, pct(r.pct_feita), pct(r.pct_sem), pct(r.pct_nao), esc(r.motivo || "—"), r.tempo != null ? r.tempo + " min" : "—"]),
        xls: D.por_tarefa.map((r) => ({ "Código": r.codigo, Tarefa: r.tarefa, Setor: r.setor, Registros: r.registros, "% feita": r.pct_feita, "% sem demanda": r.pct_sem, "% não feita": r.pct_nao, "Motivo principal": r.motivo, "Tempo médio (min)": r.tempo })) },
    };
    const A = abas[d.aba] || abas.pessoa;
    S._dashXls = () => baixarXlsx(`dashboard_${d.aba}_${ini}_a_${fim}.xlsx`, [[{ pessoa: "Por colaborador", cargo: "Por cargo", tarefa: "Por tarefa" }[d.aba], A.xls]]);
    const U = uso?.data;
    C().innerHTML = `
      <div class="titulo"><div><h1>Dashboard de produtividade</h1><p>${esc(setor ? setorNome(setor) : "Todos os setores")} · ${br(ini)} a ${br(fim)} · ${D.dias_uteis} dias úteis</p></div>
        <div class="botoes"><button class="btn" data-act="dash-xls">Exportar tabela (.xlsx)</button></div></div>
      ${U ? `<section class="cartao armaz" aria-label="Armazenamento"><b>Armazenamento</b>
        <span>Banco <span class="mbar"><i style="width:${Math.min(100, U.banco_bytes / 5242880)}%;background:var(--ok-bar)"></i></span><b>${mb(U.banco_bytes)} / 500 MB</b></span>
        <span>Fotos <span class="mbar"><i style="width:${Math.min(100, U.fotos_bytes / 10737418)}%;background:${U.fotos_bytes > 751619276 ? "#B7791F" : "var(--ok-bar)"}"></i></span><b>${mb(U.fotos_bytes)} / 1 GB</b></span>
        <span style="color:var(--sub)">${U.fotos_qtd.toLocaleString("pt-BR")} fotos · ${U.registros.toLocaleString("pt-BR")} registros</span></section>` : ""}
      <section class="cartao filtros" aria-label="Filtros">
        <div class="segm" role="group" aria-label="Período">${periodos.map(([v, t]) => `<button type="button" class="${d.periodo === v ? "on" : ""}" data-act="periodo" data-v="${v}">${t}</button>`).join("")}</div>
        ${d.periodo === "custom" ? `<label>De<input type="date" data-f="dash.ini" value="${esc(d.ini)}"></label><label>Até<input type="date" data-f="dash.fim" value="${esc(d.fim)}"></label>` : ""}
        <label>Setor<select data-f="dash.setor" ${isAdmin() ? "" : "disabled"}>${isAdmin() ? opt("", "Todos os setores", d.setor) : ""}${setoresVisiveis().map((s) => opt(s.id, s.nome, d.setor)).join("")}</select></label>
        <label>Cargo<select data-f="dash.cargo">${opt("", "Todos os cargos", d.cargo)}${cargos.map((c) => opt(c.id, (setor ? "" : setorNome(c.setor_id) + " · ") + c.nome, d.cargo)).join("")}</select></label>
        <label>Tarefa<select data-f="dash.tarefa" style="max-width:280px">${opt("", "Todas as tarefas", d.tarefa)}${tarefas.map((t) => opt(t.id, `${t.codigo} · ${t.descricao}`, d.tarefa)).join("")}</select></label>
        <label>Colaborador<select data-f="dash.pessoa">${opt("", "Todos", d.pessoa)}${pessoas.filter((p) => !d.cargo || String(p.cargo_id) === String(d.cargo)).map((p) => opt(p.id, p.nome, d.pessoa)).join("")}</select></label>
      </section>
      <section class="kpis">
        ${[["Adesão ao diário", pct(K.adesao), "dias com registro ÷ dias úteis"], ["Tarefas feitas / pessoa / dia", num(K.feitas_dia), `${K.feitas} feitas no período`],
           ["Tempo médio", K.tempo_medio != null ? K.tempo_medio + " min" : "—", "entre registros"], ["Feita · sem dem. · não feita", dist, `${totStatus} registros`],
           ["Não feitas", K.nao_feitas, "no período"], ["Evidências", pct(K.evidencia), `das “feitas” com foto${K.anexo_falhou ? ` · ${K.anexo_falhou} anexos falharam` : ""}`],
           ["Colaboradores", K.pessoas, "ativos no filtro"]].map(([a, b, c]) => `<div class="cartao kpi"><span>${a}</span><b>${b}</b><small>${c}</small></div>`).join("")}
      </section>
      <section class="linha2">
        <div class="cartao pad"><h2>Lançamentos por dia</h2>
          <div class="legenda"><span><i style="background:var(--ok-bar)"></i>Feita</span><span><i style="background:var(--sem-bar)"></i>Sem demanda</span><span><i style="background:var(--nao-bar)"></i>Não feita</span></div>
          ${D.por_dia.length ? `<div class="barras">${D.por_dia.map((x) => `<div class="coluna" title="${dm(x.data)}: ${x.feitas} feitas, ${x.sem} sem demanda, ${x.nao} não feitas"><small>${x.feitas + x.sem + x.nao}</small><div class="pilha"><div style="height:${h(x.feitas)}px;background:var(--ok-bar)"></div><div style="height:${h(x.sem)}px;background:var(--sem-bar)"></div><div style="height:${h(x.nao)}px;background:var(--nao-bar)"></div></div></div>`).join("")}</div>
          <div class="datas">${D.por_dia.map((x) => `<span>${dm(x.data)}</span>`).join("")}</div>` : `<div class="vazio">Sem registros no período.</div>`}
        </div>
        <div class="cartao pad"><h2>Motivos de “não feita”</h2>
          ${D.motivos.length ? D.motivos.map((m) => `<div class="mot"><div><span>${esc(m.motivo)}</span><b>${m.n}</b></div><div class="tr"><i style="width:${Math.round(m.n / maxMot * 100)}%"></i></div></div>`).join("") : `<div class="vazio" style="padding:0">Nenhuma não feita no período.</div>`}
        </div>
      </section>
      <section class="cartao pad">
        <div class="titulo" style="margin-bottom:12px"><div class="segm" role="tablist">${[["pessoa", "Por colaborador"], ["cargo", "Por cargo"], ["tarefa", "Por tarefa"]].map(([v, t]) => `<button type="button" role="tab" class="${d.aba === v ? "on" : ""}" data-act="aba" data-v="${v}">${t}</button>`).join("")}</div>
          <p>${{ pessoa: "Adesão abaixo de 80%, mais de 10% de não feitas ou menos de 30% de evidência geram alerta.", cargo: "Compare funções dentro e entre setores.", tarefa: "Tarefas com muito “sem demanda” podem mudar de periodicidade." }[d.aba]}</p></div>
        <div class="tabela-box"><table><thead><tr>${A.cab.map((c) => `<th>${c}</th>`).join("")}</tr></thead>
          <tbody>${A.lin.length ? A.lin.map((l) => `<tr>${l.map((c) => `<td>${c}</td>`).join("")}</tr>`).join("") : `<tr><td colspan="${A.cab.length}" class="vazio">Sem dados no filtro.</td></tr>`}</tbody></table></div>
      </section>
      <p style="margin:0;font-size:12px;color:var(--sub);line-height:1.5"><b>Adesão</b> = dias úteis com pelo menos 1 registro ÷ dias úteis do período · <b>feitas/dia</b> = registros “Feita” ÷ dias com registro · <b>tempo médio</b> = intervalo entre registros consecutivos (o 1º do dia conta desde 08:00). Feriados ficam na tabela “feriados” do Supabase.</p>`;
  }

  /* ================================================================== REGISTROS */
  function consultaRegistros(R, comContagem) {
    let q = sb.from("registros").select("id,data,criado_em,codigo,tarefa_desc,grupo_nome,status,motivo,intervalo_min,referencia,evidencia_status,evidencia_path,evidencia_kb,substitui_id,setor_id,cargo_id,user_id,profiles(nome,usuario)", comContagem ? { count: "exact" } : undefined)
      .gte("data", R.ini).lte("data", R.fim).order("criado_em", { ascending: false });
    if (R.setor) q = q.eq("setor_id", Number(R.setor));
    if (R.pessoa) q = q.eq("user_id", R.pessoa);
    if (R.status) q = q.eq("status", R.status);
    if (R.ev) q = q.eq("evidencia_status", R.ev);
    if (R.q) { const t = R.q.replace(/[%,()]/g, " ").trim(); if (t) q = q.or(`codigo.ilike.%${t}%,tarefa_desc.ilike.%${t}%`); }
    return q;
  }
  async function pRegistros() {
    const R = S.reg, POR = 50;
    const pessoas = await pessoasDoSetor(R.setor ? Number(R.setor) : null);
    const { data, count, error } = await consultaRegistros(R, true).range(R.pag * POR, R.pag * POR + POR - 1);
    if (error) { C().innerHTML = `<div class="cartao vazio">Erro: ${esc(error.message)}</div>`; return; }
    const opt = (v, t, sel) => `<option value="${esc(v)}"${String(sel) === String(v) ? " selected" : ""}>${esc(t)}</option>`;
    const paginas = Math.max(1, Math.ceil((count || 0) / POR));
    C().innerHTML = `
      <div class="titulo"><div><h1>Registros e evidências</h1><p>Uma base única com todos os lançamentos${isAdmin() ? "" : " do seu setor"}. Remarcações aparecem como registros novos; nada é apagado.</p></div>
        <div class="botoes">${isAdmin() ? `<button class="btn perigo" data-act="limpar">Limpar fotos antigas</button>` : ""}<button class="btn prim" data-act="reg-xls">Exportar planilha (.xlsx)</button></div></div>
      <section class="cartao filtros">
        <label>De<input type="date" data-f="reg.ini" value="${esc(R.ini)}"></label>
        <label>Até<input type="date" data-f="reg.fim" value="${esc(R.fim)}"></label>
        <label>Setor<select data-f="reg.setor" ${isAdmin() ? "" : "disabled"}>${isAdmin() ? opt("", "Todos", R.setor) : ""}${setoresVisiveis().map((s) => opt(s.id, s.nome, R.setor)).join("")}</select></label>
        <label>Colaborador<select data-f="reg.pessoa">${opt("", "Todos", R.pessoa)}${pessoas.map((p) => opt(p.id, p.nome, R.pessoa)).join("")}</select></label>
        <label>Status<select data-f="reg.status">${opt("", "Todos", R.status)}${Object.entries(TXT).map(([k, v]) => opt(k, v, R.status)).join("")}</select></label>
        <label>Evidência<select data-f="reg.ev">${opt("", "Todas", R.ev)}${opt("ok", "Com foto", R.ev)}${opt("falhou", "Anexo falhou", R.ev)}${opt("expirada", "Foto expirada", R.ev)}</select></label>
        <label style="flex:1 1 200px">Buscar<input type="search" data-f="reg.q" value="${esc(R.q)}" placeholder="Código ou tarefa"></label>
      </section>
      <section class="cartao" style="overflow:hidden">
        <div class="tabela-box"><table><thead><tr><th>Data</th><th>Hora</th><th>Colaborador</th><th>Setor</th><th>Código</th><th>Tarefa</th><th>Status</th><th>Motivo</th><th>Intervalo</th><th>Evidência</th></tr></thead><tbody>
        ${(data || []).map((r) => `<tr><td>${dm(r.data)}</td><td>${horaDe(r.criado_em)}</td><td><b>${esc(r.profiles?.nome || "—")}</b></td><td>${setorPill(r.setor_id)}</td>
          <td><span class="cod">${esc(r.codigo)}</span></td><td>${esc(r.tarefa_desc)}${r.substitui_id ? ' <span class="pill p-sem" title="Este registro substituiu um anterior do mesmo dia">remarcação</span>' : ""}</td>
          <td><span class="pill ${PILL[r.status]}">${TXT[r.status]}</span></td><td>${esc(r.motivo || "—")}</td><td>${r.intervalo_min ?? "—"} min</td>
          <td>${r.evidencia_status === "ok" ? `<button class="btn peq" data-act="foto" data-path="${esc(r.evidencia_path)}" data-tit="${esc(r.codigo + " · " + (r.profiles?.nome || "") + " · " + dm(r.data) + " " + horaDe(r.criado_em))}">Ver foto · ${r.evidencia_kb} KB</button>`
              : r.evidencia_status === "falhou" ? '<span class="pill p-amb">anexo falhou</span>' : r.evidencia_status === "expirada" ? '<span class="pill p-sem">expirada</span>' : "—"}</td></tr>`).join("")
          || `<tr><td colspan="10" class="vazio">Nenhum registro no filtro.</td></tr>`}
        </tbody></table></div>
        <div class="rodape"><span>${(count || 0).toLocaleString("pt-BR")} registros</span>
          <span><button class="btn peq" data-act="pag" data-v="-1" ${R.pag ? "" : "disabled"}>Anterior</button> Página ${R.pag + 1} de ${paginas} <button class="btn peq" data-act="pag" data-v="1" ${R.pag + 1 < paginas ? "" : "disabled"}>Próxima</button></span></div>
      </section>`;
  }
  async function exportarRegistros() {
    toast("Gerando a planilha…");
    const linhas = [];
    for (let i = 0; i < 200; i++) {
      const { data, error } = await consultaRegistros(S.reg, false).range(i * 1000, i * 1000 + 999);
      if (error) return toast("Erro ao exportar: " + error.message, true);
      for (const r of data) linhas.push({ Data: br(r.data), Hora: horaDe(r.criado_em), Colaborador: r.profiles?.nome, "Usuário": r.profiles?.usuario, Setor: setorNome(r.setor_id), Cargo: cargoNome(r.cargo_id),
        "Código": r.codigo, Grupo: r.grupo_nome, Tarefa: r.tarefa_desc, Status: TXT[r.status], Motivo: r.motivo || "", "Intervalo (min)": r.intervalo_min, "Referência": r.referencia,
        "Evidência": r.evidencia_status === "ok" ? "Com foto" : r.evidencia_status === "falhou" ? "Anexo falhou" : r.evidencia_status === "expirada" ? "Expirada" : "Sem foto", "Remarcação": r.substitui_id ? "Sim" : "", ID: r.id });
      if (data.length < 1000) break;
    }
    baixarXlsx(`registros_${S.reg.ini}_a_${S.reg.fim}.xlsx`, [["Registros", linhas]]);
  }

  /* ================================================================== PESSOAS */
  async function pPessoas() {
    const P = S.pes;
    let q = sb.from("profiles").select("*").order("nome");
    if (P.setor) q = q.eq("setor_id", Number(P.setor));
    if (P.perfil) q = q.eq("perfil", P.perfil);
    if (P.ativo) q = q.eq("ativo", P.ativo === "1");
    const { data, error } = await q;
    if (error) { C().innerHTML = `<div class="cartao vazio">Erro: ${esc(error.message)}</div>`; return; }
    const t = norm(P.q);
    const lista = (data || []).filter((p) => !t || norm(p.nome).includes(t) || norm(p.email).includes(t));
    S._pessoas = data || [];
    const opt = (v, x, sel) => `<option value="${esc(v)}"${String(sel) === String(v) ? " selected" : ""}>${esc(x)}</option>`;
    const temGestores = (data || []).some((p) => p.perfil === "gestor");
    C().innerHTML = `
      <div class="titulo"><div><h1>Pessoas</h1><p>Usuário = parte do e-mail antes do @. O cargo define as tarefas que aparecem no diário. ${isAdmin() ? "" : "Você vê e cadastra só o seu setor."}</p></div>
        <div class="botoes">${isAdmin() && !temGestores ? `<button class="btn" data-act="gestores-teste">Criar gestores de teste</button>` : ""}
          <button class="btn" data-act="importar-pessoas">Importar planilha</button><button class="btn prim" data-act="nova-pessoa">+ Nova pessoa</button></div></div>
      <section class="cartao filtros">
        <label style="flex:1 1 220px">Buscar<input type="search" data-f="pes.q" value="${esc(P.q)}" placeholder="Nome ou e-mail"></label>
        <label>Setor<select data-f="pes.setor" ${isAdmin() ? "" : "disabled"}>${isAdmin() ? opt("", "Todos", P.setor) : ""}${setoresVisiveis().map((s) => opt(s.id, s.nome, P.setor)).join("")}</select></label>
        <label>Perfil<select data-f="pes.perfil">${opt("", "Todos", P.perfil)}${opt("colaborador", "Colaborador", P.perfil)}${opt("gestor", "Gestor", P.perfil)}${opt("admin", "Admin", P.perfil)}</select></label>
        <label>Situação<select data-f="pes.ativo">${opt("1", "Ativos", P.ativo)}${opt("0", "Inativos", P.ativo)}${opt("", "Todos", P.ativo)}</select></label>
      </section>
      <section class="cartao" style="overflow:hidden"><div class="tabela-box"><table>
        <thead><tr><th>Nome</th><th>Usuário</th><th>Setor</th><th>Cargo</th><th>Perfil</th><th>Situação</th><th>Ações</th></tr></thead><tbody>
        ${lista.map((p) => `<tr class="${p.ativo ? "" : "inativo"}"><td><b>${esc(p.nome)}</b><br><small style="color:var(--sub)">${esc(p.email)}</small></td><td class="mono">${esc(p.usuario)}</td>
          <td>${setorPill(p.setor_id)}</td><td>${p.cargo_id ? esc(cargoNome(p.cargo_id)) : "—"}</td><td><span class="pill p-${p.perfil}">${esc(p.perfil[0].toUpperCase() + p.perfil.slice(1))}</span></td>
          <td>${p.ativo ? (p.must_change_password ? '<span class="pill p-amb">aguardando 1º acesso</span>' : '<span class="pill p-ok">ativo</span>') : '<span class="pill p-sem">inativo</span>'}</td>
          <td style="white-space:nowrap">${(isAdmin() || p.perfil !== "admin") ? `<button class="btn peq" data-act="editar-pessoa" data-id="${p.id}">Editar</button>` : ""}
            ${isAdmin() && p.id !== S.eu.id ? ` <button class="btn peq" data-act="senha" data-id="${p.id}">Gerar senha temporária</button>` : ""}</td></tr>`).join("")
          || `<tr><td colspan="7" class="vazio">Ninguém encontrado.</td></tr>`}
        </tbody></table></div><div class="rodape"><span>${lista.length} pessoas</span></div></section>`;
  }

  function formPessoa(p) {
    const novo = !p; p = p || { perfil: "colaborador", setor_id: isAdmin() ? "" : S.eu.setor_id, cargo_id: "", ativo: true };
    const perfis = isAdmin() ? ["colaborador", "gestor", "admin"] : ["colaborador", "gestor"];
    const setores = setoresVisiveis();
    modal(`<form id="f-pessoa" data-id="${novo ? "" : p.id}">
      <div class="topo-m"><h2>${novo ? "Nova pessoa" : "Editar " + esc(p.nome)}</h2><button type="button" class="x" data-fechar aria-label="Fechar">×</button></div>
      <div class="grade" style="margin-top:14px">
        <label class="campo" style="grid-column:1/-1">E-mail corporativo<input name="email" type="email" required value="${esc(p.email || "")}" placeholder="nome.sobrenome@${esc(CFG.EMAIL_DOMINIO)}"></label>
        <label class="campo" style="grid-column:1/-1">Nome<input name="nome" required value="${esc(p.nome || "")}"></label>
        <label class="campo">Setor<select name="setor_id" required ${isAdmin() ? "" : "disabled"}><option value="">Escolha…</option>${setores.map((s) => `<option value="${s.id}"${String(p.setor_id) === String(s.id) ? " selected" : ""}>${esc(s.nome)}</option>`).join("")}</select></label>
        <label class="campo">Perfil<select name="perfil">${perfis.map((x) => `<option value="${x}"${p.perfil === x ? " selected" : ""}>${x[0].toUpperCase() + x.slice(1)}</option>`).join("")}</select></label>
        <label class="campo" id="campo-cargo" style="grid-column:1/-1">Cargo (define as tarefas do diário)<select name="cargo_id"></select></label>
        ${novo ? "" : `<label class="campo" style="flex-direction:row;align-items:center;gap:8px"><input type="checkbox" name="ativo" ${p.ativo ? "checked" : ""} style="height:auto"> Ativo (desmarcar bloqueia o acesso)</label>`}
      </div>
      <p id="usr" style="margin-top:10px"></p>
      <div class="erro" id="f-erro"></div>
      <div class="fim"><button type="button" class="btn" data-fechar>Cancelar</button><button type="submit" class="btn prim">${novo ? "Salvar e gerar senha temporária" : "Salvar"}</button></div></form>`);
    const f = $("#f-pessoa");
    const sync = () => {
      const s = f.setor_id.value, perfil = f.perfil.value, atual = f.cargo_id.value || p.cargo_id;
      $("#campo-cargo").style.display = perfil === "colaborador" ? "" : "none";
      f.cargo_id.innerHTML = `<option value="">Escolha…</option>` + S.cargos.filter((c) => String(c.setor_id) === String(s) && c.ativo).map((c) => `<option value="${c.id}"${String(atual) === String(c.id) ? " selected" : ""}>${esc(c.nome)}</option>`).join("");
      $("#usr").innerHTML = f.email.value.includes("@") ? `Usuário de login: <b class="mono">${esc(f.email.value.split("@")[0].toLowerCase())}</b>` : "";
    };
    f.addEventListener("change", sync); f.email.addEventListener("input", sync); sync();
  }
  async function salvarPessoa(f) {
    const id = f.dataset.id, d = { email: f.email.value, nome: f.nome.value, setor_id: Number(f.setor_id.value) || null, perfil: f.perfil.value, cargo_id: Number(f.cargo_id.value) || null };
    if (d.perfil === "colaborador" && !d.cargo_id) { $("#f-erro").textContent = "Escolha o cargo do colaborador."; return; }
    f.querySelector("button[type=submit]").disabled = true;
    try {
      if (!id) { const r = await funcao("criar", d); cachePessoas = {}; mostrarSenha(r.usuario, r.senha, true); }
      else { await funcao("atualizar", { ...d, id, ativo: f.ativo.checked }); fechar(); toast("Cadastro atualizado."); }
      if (S.rota === "pessoas") pPessoas();
    } catch (e) { $("#f-erro").textContent = e.message; f.querySelector("button[type=submit]").disabled = false; }
  }
  function mostrarSenha(usuario, senha, nova) {
    modal(`<div class="topo-m"><h2>Senha temporária de ${esc(usuario)}</h2><button type="button" class="x" data-fechar aria-label="Fechar">×</button></div>
      <div class="senha" id="senha-tmp">${esc(senha)}</div>
      <p>Ela aparece só agora. Envie para a pessoa junto com o endereço do diário${nova ? "" : ""}: <b>${esc(CFG.DIARIO_URL)}</b>. No primeiro login ela cria uma senha própria, que ninguém mais vê.</p>
      <div class="fim"><button class="btn" data-act="copiar" data-v="${esc(`Usuário: ${usuario}\nSenha temporária: ${senha}\nAcesso: ${CFG.DIARIO_URL}`)}">Copiar mensagem</button><button class="btn prim" data-fechar>Pronto</button></div>`);
  }

  function importarPessoas() {
    modal(`<div class="topo-m"><h2>Importar pessoas de uma planilha</h2><button type="button" class="x" data-fechar aria-label="Fechar">×</button></div>
      <p>Colunas: <b>E-mail, Nome, Setor, Cargo, Perfil</b> (perfil vazio = colaborador). Quem já existe (mesmo e-mail) é atualizado, sem trocar a senha. Ao final você baixa uma planilha com o usuário e a senha temporária de cada pessoa nova.</p>
      <div class="fim" style="justify-content:flex-start"><button class="btn" data-act="modelo-pessoas">Baixar modelo</button></div>
      <label class="campo">Arquivo (.xlsx ou .csv)<input type="file" id="arq-imp" accept=".xlsx,.xls,.csv"></label>
      <label class="campo" style="flex-direction:row;align-items:center;gap:8px;font-weight:400"><input type="checkbox" id="criar-cargos" style="height:auto"> Criar cargos que ainda não existem</label>
      <div id="imp-prev"></div>
      <div class="fim"><button class="btn" data-fechar>Cancelar</button><button class="btn prim" id="btn-imp" data-act="rodar-importacao" disabled>Importar</button></div>`, true);
    $("#arq-imp").addEventListener("change", async (e) => {
      try {
        const linhas = (await lerPlanilha(e.target.files[0])).map((l) => ({ email: l.email, nome: l.nome, setor: l.setor || (isAdmin() ? "" : setorNome(S.eu.setor_id)), cargo: l.cargo, perfil: l.perfil || "colaborador" })).filter((l) => l.email);
        S._imp = linhas;
        const setoresDesc = [...new Set(linhas.map((l) => l.setor).filter((s) => !S.setores.some((x) => norm(x.nome) === norm(s))))];
        $("#imp-prev").innerHTML = `<div class="resumo-imp"><div class="p-ok"><b>${linhas.length}</b>linhas com e-mail</div><div class="p-gestor"><b>${new Set(linhas.map((l) => norm(l.setor))).size}</b>setores</div><div class="${setoresDesc.length ? "p-nao" : "p-sem"}"><b>${setoresDesc.length}</b>setores desconhecidos${setoresDesc.length ? ": " + esc(setoresDesc.join(", ")) : ""}</div></div>`;
        $("#btn-imp").disabled = !linhas.length; $("#btn-imp").textContent = `Importar ${linhas.length}`;
      } catch (err) { $("#imp-prev").innerHTML = `<p class="erro">Não consegui ler o arquivo: ${esc(err.message)}</p>`; }
    });
  }
  async function rodarImportacao(linhas, criarCargos, titulo) {
    const res = [];
    for (let i = 0; i < linhas.length; i += 40) {
      $("#imp-prev") && ($("#imp-prev").innerHTML = `<p>Importando ${Math.min(i + 40, linhas.length)} de ${linhas.length}…</p>`);
      try { const r = await funcao("importar", { linhas: linhas.slice(i, i + 40), criarCargos }); res.push(...r.resultados.map((x) => ({ ...x, linha: x.linha + i }))); }
      catch (e) { res.push({ linha: i + 2, email: "", resultado: "erro: " + e.message }); }
    }
    cachePessoas = {};
    const c = (t) => res.filter((r) => r.resultado.startsWith(t)).length;
    const planilha = res.map((r) => ({ Linha: r.linha, "E-mail": r.email, "Usuário": r.usuario, "Senha temporária": r.senha, Resultado: r.resultado, Acesso: CFG.DIARIO_URL }));
    S._impRes = planilha;
    modal(`<div class="topo-m"><h2>${esc(titulo || "Importação concluída")}</h2><button type="button" class="x" data-fechar aria-label="Fechar">×</button></div>
      <div class="resumo-imp"><div class="p-ok"><b>${c("criado")}</b>criadas</div><div class="p-gestor"><b>${c("atualizado")}</b>atualizadas</div><div class="${c("erro") ? "p-nao" : "p-sem"}"><b>${c("erro")}</b>com erro</div></div>
      ${c("erro") ? `<div class="tabela-box" style="max-height:220px"><table><tbody>${res.filter((r) => r.resultado.startsWith("erro")).map((r) => `<tr><td>linha ${r.linha}</td><td>${esc(r.email)}</td><td>${esc(r.resultado)}</td></tr>`).join("")}</tbody></table></div>` : ""}
      <p><b>Baixe agora</b> a planilha com os usuários e senhas temporárias: as senhas não ficam guardadas em lugar nenhum.</p>
      <div class="fim"><button class="btn prim" data-act="baixar-senhas">Baixar usuários e senhas (.xlsx)</button><button class="btn" data-fechar>Fechar</button></div>`, true);
    if (S.rota === "pessoas") pPessoas();
  }

  /* ================================================================== ESTRUTURA */
  async function pEstrutura() {
    const E = S.est, vis = setoresVisiveis();
    if (!E.setor || !vis.some((s) => s.id === E.setor)) E.setor = vis[0]?.id || null;
    if (!E.setor) { C().innerHTML = `<div class="cartao vazio">Nenhum setor cadastrado.</div>`; return; }
    const pode = isAdmin() || S.eu.setor_id === E.setor;
    const [cg, gr, tf, pes] = await Promise.all([
      sb.from("cargos").select("*").eq("setor_id", E.setor).order("nome"),
      sb.from("grupos").select("*").eq("setor_id", E.setor).order("letra"),
      sb.from("tarefas").select("*, grupos(letra,nome), tarefa_cargos(cargo_id)").eq("setor_id", E.setor).order("codigo"),
      sb.from("profiles").select("cargo_id").eq("setor_id", E.setor).eq("perfil", "colaborador").eq("ativo", true),
    ]);
    const cargos = cg.data || [], grupos = gr.data || [], tarefas = tf.data || [];
    S.cargos = [...S.cargos.filter((c) => c.setor_id !== E.setor), ...cargos];
    S._est = { cargos, grupos, tarefas };
    if (E.cargo !== "todas" && !cargos.some((c) => c.id === E.cargo)) E.cargo = cargos[0]?.id || "todas";
    const nPes = (id) => (pes.data || []).filter((p) => p.cargo_id === id).length;
    const lista = E.cargo === "todas" ? tarefas : tarefas.filter((t) => t.tarefa_cargos.some((x) => x.cargo_id === E.cargo));
    const nomeCargo = E.cargo === "todas" ? "todas as tarefas do setor" : cargos.find((c) => c.id === E.cargo)?.nome;
    const cor = setorCor(E.setor);
    C().innerHTML = `
      <div class="titulo"><div><h1>Setores, cargos e tarefas</h1><p>Setor → cargo → tarefas. ${isAdmin() ? "Como admin você edita todos os setores." : "Você edita o seu setor."}</p></div>
        <div class="botoes">${pode ? `<button class="btn" data-act="importar-tarefas">Importar planilha de tarefas</button>` : ""}${isAdmin() ? `<button class="btn" data-act="novo-setor">+ Novo setor</button>` : ""}</div></div>
      <section class="cartao setores" aria-label="Setores">${vis.map((s) => `<button type="button" class="${s.id === E.setor ? "on" : ""}" style="${s.id === E.setor ? `background:${esc(s.cor)};border-color:${esc(s.cor)}` : ""}" data-act="setor" data-id="${s.id}"><i style="background:${s.id === E.setor ? "#fff" : esc(s.cor)}"></i>${esc(s.nome)}</button>`).join("")}</section>
      <div class="estr">
        <aside class="cartao">
          <h3>CARGOS · ${esc(setorNome(E.setor).toUpperCase())}</h3>
          ${cargos.map((c) => `<button type="button" class="it ${E.cargo === c.id ? "on" : ""}" data-act="cargo" data-id="${c.id}" style="${c.ativo ? "" : "opacity:.55"}"><span>${esc(c.nome)}</span><small>${tarefas.filter((t) => t.tarefa_cargos.some((x) => x.cargo_id === c.id)).length} tarefas · ${nPes(c.id)} pessoas</small></button>`).join("") || `<p class="vazio" style="padding:4px">Nenhum cargo ainda.</p>`}
          <button type="button" class="it ${E.cargo === "todas" ? "on" : ""}" data-act="cargo" data-id="todas"><span>Todas as tarefas do setor</span><small>${tarefas.length}</small></button>
          ${pode ? `<form id="f-cargo"><input name="nome" placeholder="Novo cargo" required><button class="btn prim peq" type="submit">Criar</button></form>` : ""}
          ${pode && E.cargo !== "todas" && E.cargo ? `<div style="display:flex;gap:6px;padding-top:6px"><button class="btn peq" data-act="renomear-cargo">Renomear</button><button class="btn peq" data-act="ativar-cargo">${cargos.find((c) => c.id === E.cargo)?.ativo ? "Desativar" : "Ativar"}</button></div>` : ""}
          <h3 style="margin-top:14px">GRUPOS</h3>
          ${grupos.map((g) => `<div class="it" style="cursor:default"><span><b style="color:${esc(cor)}">${esc(g.letra)}</b> · ${esc(g.nome)}</span>${pode ? `<button class="btn peq" data-act="renomear-grupo" data-id="${g.id}">Renomear</button>` : ""}</div>`).join("") || `<p class="vazio" style="padding:4px">Nenhum grupo ainda.</p>`}
          ${pode ? `<form id="f-grupo"><input name="nome" placeholder="Novo grupo (ex.: Rotina diária)" required><button class="btn prim peq" type="submit">Criar</button></form>` : ""}
        </aside>
        <section>
          <div class="cartao" style="overflow:hidden"><div class="pad" style="padding-bottom:8px"><h2 style="margin:0">Tarefas · ${esc(nomeCargo || "—")}</h2></div>
            <div class="tabela-box"><table><thead><tr><th>Código</th><th>Grupo</th><th>Tarefa</th><th>Cargos</th><th>Periodicidade</th><th>Pedir evidência</th><th>Ativa</th><th></th></tr></thead><tbody>
            ${lista.map((t) => `<tr class="${t.ativa ? "" : "inativo"}"><td><span class="cod">${esc(t.codigo)}</span></td><td>${esc(t.grupos?.nome || "")}</td><td>${esc(t.descricao)}</td>
              <td style="font-size:12px;color:var(--sub)">${t.tarefa_cargos.map((x) => esc(cargos.find((c) => c.id === x.cargo_id)?.nome || "")).join(", ") || "—"}</td>
              <td>${esc(t.periodicidade)}</td><td>${t.pedir_evidencia ? "Sim" : "Não"}</td>
              <td>${pode ? `<button class="chave ${t.ativa ? "on" : ""}" data-act="ativar-tarefa" data-id="${t.id}" aria-pressed="${t.ativa}" aria-label="Ativar ou desativar"><span></span></button>` : (t.ativa ? "Sim" : "Não")}</td>
              <td>${pode ? `<button class="btn peq" data-act="editar-tarefa" data-id="${t.id}">Editar</button>` : ""}</td></tr>`).join("") || `<tr><td colspan="8" class="vazio">Nenhuma tarefa aqui ainda.</td></tr>`}
            </tbody></table></div></div>
          ${pode ? `<div class="cartao pad">${formTarefaHtml(null, grupos, cargos)}</div>` : ""}
        </section>
      </div>`;
  }
  function formTarefaHtml(t, grupos, cargos) {
    const sel = t ? t.tarefa_cargos.map((x) => x.cargo_id) : (S.est.cargo && S.est.cargo !== "todas" ? [S.est.cargo] : []);
    return `<form id="f-tarefa" data-id="${t ? t.id : ""}">
      <h2>${t ? "Editar " + esc(t.codigo) : "Nova tarefa"}</h2>
      <div class="grade">
        <label class="campo" style="grid-column:1/-1">Descrição<input name="descricao" required value="${esc(t?.descricao || "")}"></label>
        <label class="campo">Grupo<select name="grupo_id" required ${t ? "disabled" : ""}><option value="">Escolha…</option>${grupos.map((g) => `<option value="${g.id}"${t?.grupo_id === g.id ? " selected" : ""}>${esc(g.letra)} · ${esc(g.nome)}</option>`).join("")}</select></label>
        <label class="campo">Periodicidade esperada<select name="periodicidade">${PER.map((p) => `<option${(t?.periodicidade || "Diária") === p ? " selected" : ""}>${p}</option>`).join("")}</select></label>
        <label class="campo">Evidência<select name="pedir_evidencia"><option value="1"${!t || t.pedir_evidencia ? " selected" : ""}>Pedir (pode enviar sem)</option><option value="0"${t && !t.pedir_evidencia ? " selected" : ""}>Não pedir</option></select></label>
      </div>
      <p style="margin:12px 0 6px;font-size:13px;font-weight:600">Aparece para os cargos</p>
      <div class="cheques">${cargos.map((c) => `<label><input type="checkbox" name="cargo" value="${c.id}"${sel.includes(c.id) ? " checked" : ""}>${esc(c.nome)}</label>`).join("") || '<span style="color:var(--sub);font-size:13px">Crie um cargo primeiro.</span>'}</div>
      <div class="erro" id="t-erro"></div>
      <div class="fim">${t ? '<button type="button" class="btn" data-fechar>Cancelar</button>' : ""}<button type="submit" class="btn prim">${t ? "Salvar" : "Adicionar tarefa"}</button></div>
      ${t ? "" : `<p style="margin-top:6px;font-size:12px;color:var(--sub)">O código (ex.: A007) é gerado sozinho a partir da letra do grupo.</p>`}</form>`;
  }
  async function salvarTarefa(f) {
    const id = Number(f.dataset.id) || null;
    const cargos = [...f.querySelectorAll("input[name=cargo]:checked")].map((x) => Number(x.value));
    if (!cargos.length) { $("#t-erro").textContent = "Marque pelo menos um cargo."; return; }
    const dados = { descricao: f.descricao.value.trim(), periodicidade: f.periodicidade.value, pedir_evidencia: f.pedir_evidencia.value === "1" };
    let tid = id;
    if (id) {
      const { error } = await sb.from("tarefas").update(dados).eq("id", id);
      if (error) { $("#t-erro").textContent = error.message; return; }
      const atuais = S._est.tarefas.find((t) => t.id === id).tarefa_cargos.map((x) => x.cargo_id);
      const tirar = atuais.filter((c) => !cargos.includes(c)), por = cargos.filter((c) => !atuais.includes(c));
      if (tirar.length) await sb.from("tarefa_cargos").delete().eq("tarefa_id", id).in("cargo_id", tirar);
      if (por.length) await sb.from("tarefa_cargos").insert(por.map((c) => ({ tarefa_id: id, cargo_id: c })));
      fechar();
    } else {
      const { data, error } = await sb.from("tarefas").insert({ ...dados, grupo_id: Number(f.grupo_id.value), setor_id: S.est.setor }).select("id,codigo").single();
      if (error) { $("#t-erro").textContent = error.message; return; }
      tid = data.id;
      const r = await sb.from("tarefa_cargos").insert(cargos.map((c) => ({ tarefa_id: tid, cargo_id: c })));
      if (r.error) { $("#t-erro").textContent = r.error.message; return; }
      toast(`Tarefa ${data.codigo} criada.`);
    }
    cacheTarefas = {};
    pEstrutura();
  }
  function importarTarefas() {
    modal(`<div class="topo-m"><h2>Importar tarefas para ${esc(setorNome(S.est.setor))}</h2><button type="button" class="x" data-fechar aria-label="Fechar">×</button></div>
      <p>Colunas: <b>Cargo, Grupo, Tarefa, Periodicidade, Evidência</b>. Uma linha por tarefa e cargo; a mesma tarefa em vários cargos pode repetir com cargos diferentes. Cargos e grupos que não existem são criados. Periodicidade: Diária, Semanal, Mensal ou Eventual. Evidência: Sim ou Não.</p>
      <div class="fim" style="justify-content:flex-start"><button class="btn" data-act="modelo-tarefas">Baixar modelo</button></div>
      <label class="campo">Arquivo (.xlsx ou .csv)<input type="file" id="arq-tar" accept=".xlsx,.xls,.csv"></label>
      <div id="tar-prev"></div>
      <div class="fim"><button class="btn" data-fechar>Cancelar</button><button class="btn prim" id="btn-tar" data-act="rodar-tarefas" disabled>Importar</button></div>`, true);
    $("#arq-tar").addEventListener("change", async (e) => {
      try {
        const linhas = (await lerPlanilha(e.target.files[0])).map((l) => ({ cargo: l.cargo, grupo: l.grupo, tarefa: l.tarefa || l.descricao, per: l.periodicidade, ev: l.evidencia })).filter((l) => l.cargo && l.grupo && l.tarefa);
        S._tar = linhas;
        $("#tar-prev").innerHTML = `<div class="resumo-imp"><div class="p-ok"><b>${linhas.length}</b>linhas válidas</div><div class="p-gestor"><b>${new Set(linhas.map((l) => norm(l.cargo))).size}</b>cargos</div><div class="p-sem"><b>${new Set(linhas.map((l) => norm(l.tarefa))).size}</b>tarefas distintas</div></div>`;
        $("#btn-tar").disabled = !linhas.length;
      } catch (err) { $("#tar-prev").innerHTML = `<p class="erro">Não consegui ler o arquivo: ${esc(err.message)}</p>`; }
    });
  }
  async function rodarTarefas() {
    const setor = S.est.setor, L = S._tar || [];
    $("#tar-prev").innerHTML = "<p>Importando…</p>";
    let { cargos, grupos, tarefas } = S._est;
    cargos = [...cargos]; grupos = [...grupos]; tarefas = [...tarefas];
    let criadas = 0, vinculos = 0; const erros = [];
    for (const [i, l] of L.entries()) {
      try {
        let c = cargos.find((x) => norm(x.nome) === norm(l.cargo));
        if (!c) { const r = await sb.from("cargos").insert({ setor_id: setor, nome: l.cargo }).select().single(); if (r.error) throw r.error; c = r.data; cargos.push(c); }
        let g = grupos.find((x) => norm(x.nome) === norm(l.grupo));
        if (!g) { const r = await sb.from("grupos").insert({ setor_id: setor, nome: l.grupo }).select().single(); if (r.error) throw r.error; g = r.data; grupos.push(g); }
        let t = tarefas.find((x) => norm(x.descricao) === norm(l.tarefa));
        if (!t) {
          const per = PER.find((p) => norm(p) === norm(l.per)) || (norm(l.per).startsWith("diar") ? "Diária" : "Diária");
          const r = await sb.from("tarefas").insert({ setor_id: setor, grupo_id: g.id, descricao: l.tarefa, periodicidade: per, pedir_evidencia: !norm(l.ev).startsWith("n") }).select("*").single();
          if (r.error) throw r.error; t = { ...r.data, tarefa_cargos: [] }; tarefas.push(t); criadas++;
        }
        if (!t.tarefa_cargos.some((x) => x.cargo_id === c.id)) {
          const r = await sb.from("tarefa_cargos").insert({ tarefa_id: t.id, cargo_id: c.id }); if (r.error) throw r.error;
          t.tarefa_cargos.push({ cargo_id: c.id }); vinculos++;
        }
      } catch (e) { erros.push(`linha ${i + 2}: ${e.message}`); }
    }
    cacheTarefas = {};
    const { data: cg } = await sb.from("cargos").select("*").order("nome"); S.cargos = cg || S.cargos;
    modal(`<div class="topo-m"><h2>Importação de tarefas concluída</h2><button type="button" class="x" data-fechar aria-label="Fechar">×</button></div>
      <div class="resumo-imp"><div class="p-ok"><b>${criadas}</b>tarefas novas</div><div class="p-gestor"><b>${vinculos}</b>vínculos com cargos</div><div class="${erros.length ? "p-nao" : "p-sem"}"><b>${erros.length}</b>erros</div></div>
      ${erros.length ? `<p class="erro">${erros.slice(0, 20).map(esc).join("<br>")}</p>` : ""}
      <div class="fim"><button class="btn prim" data-fechar>Fechar</button></div>`);
    pEstrutura();
  }
  function pedirTexto(titulo, valor, cb) {
    modal(`<form id="f-texto"><div class="topo-m"><h2>${esc(titulo)}</h2><button type="button" class="x" data-fechar aria-label="Fechar">×</button></div>
      <label class="campo" style="margin-top:12px">Nome<input name="v" required value="${esc(valor || "")}"></label><div class="erro" id="x-erro"></div>
      <div class="fim"><button type="button" class="btn" data-fechar>Cancelar</button><button class="btn prim" type="submit">Salvar</button></div></form>`);
    S._texto = cb;
  }

  /* ------------------------------------------------------------ eventos */
  document.addEventListener("change", (e) => {
    const k = e.target.dataset.f; if (!k) return;
    const [grupo, campo] = k.split(".");
    S[grupo][campo] = e.target.value;
    if (grupo === "dash") { if (campo === "setor") { S.dash.cargo = ""; S.dash.tarefa = ""; S.dash.pessoa = ""; } if (campo === "cargo") S.dash.pessoa = ""; pDashboard(); }
    if (grupo === "reg") { S.reg.pag = 0; if (campo === "setor") S.reg.pessoa = ""; pRegistros(); }
    if (grupo === "pes") pPessoas();
  });
  let tBusca;
  document.addEventListener("input", (e) => {
    const k = e.target.dataset.f; if (!k || e.target.type !== "search") return;
    clearTimeout(tBusca);
    tBusca = setTimeout(() => { const [g, c] = k.split("."); S[g][c] = e.target.value; if (g === "reg") { S.reg.pag = 0; pRegistros(); } else pPessoas(); }, 400);
  });

  document.addEventListener("submit", async (e) => {
    e.preventDefault();
    const f = e.target;
    if (f.id === "f-login") {
      const u = f.u.value.trim().toLowerCase(), email = u.includes("@") ? u : `${u}@${CFG.EMAIL_DOMINIO}`;
      const { data, error } = await sb.auth.signInWithPassword({ email, password: f.s.value });
      if (error) { S.msg = "Usuário ou senha inválidos."; $("#app").innerHTML = telaLogin(); return; }
      S.msg = ""; await iniciarSessao(data.user);
    } else if (f.id === "f-senha") {
      if (f.a.value.length < 8 || f.a.value !== f.b.value) { S.msg = f.a.value.length < 8 ? "Use pelo menos 8 caracteres." : "As senhas não conferem."; $("#app").innerHTML = telaSenha(); return; }
      const { error } = await sb.auth.updateUser({ password: f.a.value });
      if (error) { S.msg = error.message; $("#app").innerHTML = telaSenha(); return; }
      await sb.rpc("senha_trocada"); S.msg = ""; await iniciarSessao(S.user);
    } else if (f.id === "f-pessoa") salvarPessoa(f);
    else if (f.id === "f-tarefa") salvarTarefa(f);
    else if (f.id === "f-cargo") {
      const r = await sb.from("cargos").insert({ setor_id: S.est.setor, nome: f.nome.value.trim() }).select().single();
      if (r.error) return toast(r.error.message.includes("duplicate") ? "Esse cargo já existe." : r.error.message, true);
      S.cargos.push(r.data); S.est.cargo = r.data.id; pEstrutura();
    } else if (f.id === "f-grupo") {
      const r = await sb.from("grupos").insert({ setor_id: S.est.setor, nome: f.nome.value.trim() });
      if (r.error) return toast(r.error.message.includes("duplicate") ? "Esse grupo já existe." : r.error.message, true);
      pEstrutura();
    } else if (f.id === "f-texto") {
      try { await S._texto(f.v.value.trim(), f); } catch (err) { $("#x-erro").textContent = err.message; }
    } else if (f.id === "f-setor") {
      const r = await sb.from("setores").insert({ nome: f.nome.value.trim(), cor: f.cor.value });
      if (r.error) { $("#x-erro").textContent = r.error.message; return; }
      fechar(); await carregarBase(); pEstrutura();
    } else if (f.id === "f-limpar") {
      const dias = Number(f.dias.value);
      f.querySelector("button[type=submit]").disabled = true;
      try { const r = await funcao("limpar_evidencias", { dias }); fechar(); toast(`${r.removidas} fotos com mais de ${r.dias} dias foram removidas.`); pRegistros(); }
      catch (err) { $("#x-erro").textContent = err.message; f.querySelector("button[type=submit]").disabled = false; }
    }
  });

  document.addEventListener("click", async (e) => {
    const b = e.target.closest("[data-act]"); if (!b) return;
    const a = b.dataset.act;
    try {
      switch (a) {
        case "sair": await sb.auth.signOut(); S.eu = null; $("#app").innerHTML = telaLogin(); break;
        case "periodo": S.dash.periodo = b.dataset.v; if (b.dataset.v === "custom" && !S.dash.ini) { S.dash.ini = addDias(hojeISO(), -29); S.dash.fim = hojeISO(); } pDashboard(); break;
        case "aba": S.dash.aba = b.dataset.v; pDashboard(); break;
        case "dash-xls": S._dashXls && S._dashXls(); break;
        case "pag": S.reg.pag += Number(b.dataset.v); pRegistros(); break;
        case "reg-xls": exportarRegistros(); break;
        case "foto": {
          const { data, error } = await sb.storage.from("evidencias").createSignedUrl(b.dataset.path, 600);
          if (error) return toast("Não consegui abrir a foto: " + error.message, true);
          modal(`<div class="topo-m"><h2>${esc(b.dataset.tit)}</h2><button type="button" class="x" data-fechar aria-label="Fechar">×</button></div>
            <img class="foto" src="${esc(data.signedUrl)}" alt="Evidência ${esc(b.dataset.tit)}">
            <div class="fim"><a class="btn" href="${esc(data.signedUrl)}" target="_blank" rel="noopener" style="display:inline-flex;align-items:center;text-decoration:none">Abrir em nova aba</a><button class="btn prim" data-fechar>Fechar</button></div>`, true);
          break;
        }
        case "limpar":
          modal(`<form id="f-limpar"><div class="topo-m"><h2>Limpar fotos antigas</h2><button type="button" class="x" data-fechar aria-label="Fechar">×</button></div>
            <p style="margin:12px 0">Apaga do armazenamento as fotos mais antigas que o prazo escolhido. Os registros continuam; a evidência passa a aparecer como “expirada”. Se quiser guardar as fotos, baixe antes pelo painel do Supabase (Storage › evidencias).</p>
            <label class="campo">Apagar fotos com mais de<select name="dias"><option value="60">60 dias</option><option value="90" selected>90 dias</option><option value="180">180 dias</option><option value="365">365 dias</option></select></label>
            <div class="erro" id="x-erro"></div><div class="fim"><button type="button" class="btn" data-fechar>Cancelar</button><button type="submit" class="btn perigo">Apagar fotos</button></div></form>`);
          break;
        case "nova-pessoa": formPessoa(null); break;
        case "editar-pessoa": formPessoa(S._pessoas.find((p) => p.id === b.dataset.id)); break;
        case "senha": {
          const p = S._pessoas.find((x) => x.id === b.dataset.id);
          modal(`<div class="topo-m"><h2>Gerar senha temporária</h2><button type="button" class="x" data-fechar aria-label="Fechar">×</button></div>
            <p>A senha atual de <b>${esc(p.nome)}</b> deixa de funcionar. A pessoa entra com a temporária e cria uma nova no primeiro acesso.</p>
            <div class="fim"><button class="btn" data-fechar>Cancelar</button><button class="btn prim" data-act="senha-ok" data-id="${p.id}">Gerar senha</button></div>`);
          break;
        }
        case "senha-ok": { const r = await funcao("resetar_senha", { id: b.dataset.id }); mostrarSenha(r.usuario, r.senha, false); pPessoas(); break; }
        case "copiar": await navigator.clipboard.writeText(b.dataset.v); toast("Mensagem copiada."); break;
        case "importar-pessoas": importarPessoas(); break;
        case "modelo-pessoas": {
          const ex = setoresVisiveis()[0], cg = S.cargos.find((c) => c.setor_id === ex?.id);
          baixarXlsx("modelo_pessoas.xlsx", [["Pessoas", [{ "E-mail": `nome.sobrenome@${CFG.EMAIL_DOMINIO}`, Nome: "Nome Sobrenome", Setor: ex?.nome || "", Cargo: cg?.nome || "", Perfil: "colaborador" }]],
            ["Setores e cargos válidos", S.cargos.filter((c) => setoresVisiveis().some((s) => s.id === c.setor_id)).map((c) => ({ Setor: setorNome(c.setor_id), Cargo: c.nome }))]]);
          break;
        }
        case "rodar-importacao": b.disabled = true; await rodarImportacao(S._imp || [], $("#criar-cargos")?.checked); break;
        case "baixar-senhas": baixarXlsx(`usuarios_e_senhas_${hojeISO()}.xlsx`, [["Usuários", S._impRes || []]]); break;
        case "gestores-teste": {
          const slug = (s) => norm(s).replace(/[^a-z0-9]+/g, "");
          const linhas = S.setores.map((s) => ({ email: `gestor.${slug(s.nome)}@${CFG.EMAIL_DOMINIO}`, nome: `Gestor ${s.nome} (teste)`, setor: s.nome, cargo: "", perfil: "gestor" }));
          modal(`<div class="topo-m"><h2>Criar gestores de teste</h2><button type="button" class="x" data-fechar aria-label="Fechar">×</button></div>
            <p>Cria ${linhas.length} gestores, um por setor (${esc(linhas.map((l) => l.email.split("@")[0]).join(", "))}). Depois troque e-mail e senha de cada um pelo botão Editar e por Gerar senha temporária.</p>
            <div id="imp-prev"></div><div class="fim"><button class="btn" data-fechar>Cancelar</button><button class="btn prim" data-act="gestores-ok">Criar</button></div>`);
          S._imp = linhas; break;
        }
        case "gestores-ok": b.disabled = true; await rodarImportacao(S._imp, false, "Gestores de teste criados"); break;
        case "setor": S.est.setor = Number(b.dataset.id); S.est.cargo = null; pEstrutura(); break;
        case "cargo": S.est.cargo = b.dataset.id === "todas" ? "todas" : Number(b.dataset.id); pEstrutura(); break;
        case "renomear-cargo": {
          const c = S._est.cargos.find((x) => x.id === S.est.cargo);
          pedirTexto("Renomear cargo", c.nome, async (v) => { const r = await sb.from("cargos").update({ nome: v }).eq("id", c.id); if (r.error) throw r.error; fechar(); const { data } = await sb.from("cargos").select("*").order("nome"); S.cargos = data || S.cargos; pEstrutura(); });
          break;
        }
        case "ativar-cargo": { const c = S._est.cargos.find((x) => x.id === S.est.cargo); const r = await sb.from("cargos").update({ ativo: !c.ativo }).eq("id", c.id); if (r.error) throw r.error; pEstrutura(); break; }
        case "renomear-grupo": {
          const g = S._est.grupos.find((x) => x.id === Number(b.dataset.id));
          pedirTexto("Renomear grupo " + g.letra, g.nome, async (v) => { const r = await sb.from("grupos").update({ nome: v }).eq("id", g.id); if (r.error) throw r.error; fechar(); pEstrutura(); });
          break;
        }
        case "ativar-tarefa": { const t = S._est.tarefas.find((x) => x.id === Number(b.dataset.id)); const r = await sb.from("tarefas").update({ ativa: !t.ativa }).eq("id", t.id); if (r.error) throw r.error; cacheTarefas = {}; pEstrutura(); break; }
        case "editar-tarefa": { const t = S._est.tarefas.find((x) => x.id === Number(b.dataset.id)); modal(formTarefaHtml(t, S._est.grupos, S._est.cargos), true); break; }
        case "importar-tarefas": importarTarefas(); break;
        case "rodar-tarefas": b.disabled = true; await rodarTarefas(); break;
        case "modelo-tarefas":
          baixarXlsx("modelo_tarefas.xlsx", [["Tarefas", [
            { Cargo: "Assistente Comercial", Grupo: "Rotina diária", Tarefa: "Tratativas de e-mails", Periodicidade: "Diária", "Evidência": "Sim" },
            { Cargo: "Consultor Comercial", Grupo: "Rotina diária", Tarefa: "Tratativas de e-mails", Periodicidade: "Diária", "Evidência": "Sim" },
            { Cargo: "Assistente Comercial", Grupo: "Comitês", Tarefa: "Participar do Comitê de Clientes", Periodicidade: "Semanal", "Evidência": "Não" }]]]);
          break;
        case "novo-setor":
          modal(`<form id="f-setor"><div class="topo-m"><h2>Novo setor</h2><button type="button" class="x" data-fechar aria-label="Fechar">×</button></div>
            <div class="grade" style="margin-top:12px"><label class="campo">Nome<input name="nome" required></label><label class="campo">Cor<input name="cor" type="color" value="#2C4BA8" style="height:40px;padding:2px"></label></div>
            <div class="erro" id="x-erro"></div><div class="fim"><button type="button" class="btn" data-fechar>Cancelar</button><button class="btn prim" type="submit">Criar setor</button></div></form>`);
          break;
      }
    } catch (err) { toast(err.message || String(err), true); }
  });

  /* ------------------------------------------------------------ início */
  (async () => {
    if (!CFG.SUPABASE_URL || CFG.SUPABASE_URL.includes("SEU-PROJETO")) {
      $("#app").innerHTML = telaAviso("Falta configurar", "Edite o arquivo config.js com o endereço e a chave do seu Supabase (veja o TUTORIAL.md).");
      return;
    }
    const { data } = await sb.auth.getSession();
    if (data.session) await iniciarSessao(data.session.user);
    else $("#app").innerHTML = telaLogin();
  })();
})();
