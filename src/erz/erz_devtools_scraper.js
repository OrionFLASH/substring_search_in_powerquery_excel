/**
 * ERZ DevTools scraper (erzrf.ru)
 *
 * Запуск:
 * 1. Открыть https://erzrf.ru/zastroyschiki?... в браузере
 * 2. DevTools → Console → вставить содержимое этого файла → Enter
 * 3. В модалке пройти этапы: регионы → группы → компании → скачать JSON
 *
 * Куки берутся из текущей сессии (credentials: 'include').
 */
(function () {
  'use strict';

  const COST_TYPE = '1';
  const API = '/erz-rest/api/v1';
  const ROOT_ID = 'erz-devtools-scraper-root';
  /** Сколько попыток на один запрос (включая первую). */
  const FETCH_MAX_ATTEMPTS = 3;
  /** Базовая пауза перед повтором при обычной ошибке, мс. */
  const RETRY_BASE_MS = 2000;
  /** Пауза перед повтором при 502/503/504 / таймауте / HTML gateway, мс. */
  const RETRY_GATEWAY_MS = 8000;

  /** @type {{
   *   pauseMs: number,
   *   saveIntermediate: boolean,
   *   namesMode: 'one'|'all',
   *   joinMode: 'one'|'all',
   *   regionsRaw: Array,
   *   regionsSelected: Map<string, object>,
   *   groups: Map<string, object>,
   *   groupsSelected: Set<string>,
   *   busy: boolean,
   *   abort: boolean,
   * }} */
  const state = {
    pauseMs: 400,
    saveIntermediate: false,
    namesMode: 'one',
    // join по умолчанию «все регионы»: API отдаёт компании только для region в запросе
    joinMode: 'all',
    regionsRaw: [],
    regionsSelected: new Map(),
    groups: new Map(),
    groupsSelected: new Set(),
    busy: false,
    abort: false,
  };

  // ─── утилиты ─────────────────────────────────────────────────────────────

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function ts() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    return (
      d.getFullYear() +
      p(d.getMonth() + 1) +
      p(d.getDate()) +
      '_' +
      p(d.getHours()) +
      p(d.getMinutes()) +
      p(d.getSeconds())
    );
  }

  function safeFilePart(s) {
    return String(s || 'x')
      .replace(/[^\w\-]+/g, '_')
      .slice(0, 80);
  }

  function downloadJson(filename, data) {
    const blob = new Blob([JSON.stringify(data, null, 2)], {
      type: 'application/json;charset=utf-8',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  function maybeSaveIntermediate(stage, id, urlId, data) {
    if (!state.saveIntermediate) return;
    const name =
      stage +
      '_' +
      safeFilePart(id) +
      '_' +
      safeFilePart(urlId) +
      '_' +
      ts() +
      '.json';
    downloadJson(name, data);
  }

  function isGatewayStatus(status) {
    return status === 408 || status === 429 || status === 502 || status === 503 || status === 504;
  }

  function looksLikeHtmlError(text) {
    if (!text) return false;
    const t = String(text).trim().slice(0, 200).toLowerCase();
    return (
      t.startsWith('<!doctype') ||
      t.startsWith('<html') ||
      t.indexOf('gateway time-out') >= 0 ||
      t.indexOf('504 gateway') >= 0
    );
  }

  function retryWaitMs(attempt, isGateway) {
    // attempt: номер неудачной попытки (1, 2, …)
    const base = isGateway ? RETRY_GATEWAY_MS : RETRY_BASE_MS;
    const pause = Math.max(100, Math.min(2000, Number(state.pauseMs) || 400));
    // Экспоненциальный рост: 8с → 16с (gateway) или 2с → 4с (+ учёт паузы UI)
    return Math.round(base * Math.pow(2, attempt - 1) + pause * attempt);
  }

  /**
   * GET JSON с повторами при HTTP/сетевых/парс-ошибках.
   * Успешный ответ возвращается сразу; после FETCH_MAX_ATTEMPTS неудач — throw.
   */
  async function fetchJson(pathWithQuery) {
    const url = pathWithQuery.startsWith('http')
      ? pathWithQuery
      : window.location.origin + pathWithQuery;

    let lastError = null;

    for (let attempt = 1; attempt <= FETCH_MAX_ATTEMPTS; attempt++) {
      checkAbort();
      if (attempt === 1) {
        log('→ ' + url, 'req');
      } else {
        log('↻ повтор ' + attempt + '/' + FETCH_MAX_ATTEMPTS + ': ' + url, 'req');
      }

      let isGateway = false;
      try {
        const res = await fetch(url, {
          method: 'GET',
          credentials: 'include',
          headers: {
            Accept: 'application/json, text/plain, */*',
          },
        });

        const contentType = (res.headers.get('content-type') || '').toLowerCase();
        const rawText = await res.text();
        isGateway =
          isGatewayStatus(res.status) ||
          looksLikeHtmlError(rawText) ||
          (contentType.indexOf('text/html') >= 0 && !res.ok);

        if (!res.ok) {
          throw new Error(
            'HTTP ' + res.status + ' для ' + url + (isGateway ? ' (gateway/timeout)' : '')
          );
        }

        if (looksLikeHtmlError(rawText) || contentType.indexOf('text/html') >= 0) {
          isGateway = true;
          throw new Error('Ответ HTML вместо JSON для ' + url);
        }

        let data;
        try {
          data = rawText ? JSON.parse(rawText) : null;
        } catch (parseErr) {
          isGateway = looksLikeHtmlError(rawText);
          throw new Error(
            'Некорректный JSON для ' +
              url +
              ': ' +
              String(parseErr && parseErr.message ? parseErr.message : parseErr)
          );
        }

        if (attempt > 1) {
          log('✓ повтор успешен (попытка ' + attempt + ')', 'ok');
        }
        return data;
      } catch (e) {
        lastError = e;
        const msg = String(e && e.message ? e.message : e);
        if (state.abort || /остановлено пользователем/i.test(msg)) {
          throw e instanceof Error ? e : new Error(msg);
        }
        // Сетевой сбой / abort сети тоже считаем «тяжёлым»
        if (
          !isGateway &&
          (/failed to fetch|networkerror|load failed|timeout/i.test(msg) ||
            (e && e.name === 'TypeError'))
        ) {
          isGateway = true;
        }

        if (attempt >= FETCH_MAX_ATTEMPTS) {
          break;
        }

        const waitMs = retryWaitMs(attempt, isGateway);
        log(
          '⚠ ошибка попытки ' +
            attempt +
            '/' +
            FETCH_MAX_ATTEMPTS +
            ': ' +
            msg +
            ' — жду ' +
            waitMs +
            ' мс перед повтором' +
            (isGateway ? ' (увеличенная пауза)' : ''),
          'err'
        );
        await sleep(waitMs);
        checkAbort();
      }
    }

    throw lastError || new Error('Запрос не выполнен: ' + url);
  }

  async function pause() {
    const ms = Math.max(100, Math.min(2000, Number(state.pauseMs) || 400));
    await sleep(ms);
  }

  function companyKey(c) {
    return [c.id || '', c.inn || '', c.ogrn || '', c.urlId || ''].join('|');
  }

  function upsertCompany(map, raw, location) {
    const key = companyKey(raw);
    let item = map.get(key);
    if (!item) {
      item = {
        id: String(raw.id || ''),
        inn: String(raw.inn || ''),
        name: String(raw.name || ''),
        ogrn: String(raw.ogrn || ''),
        urlId: String(raw.urlId || ''),
        locations: [],
      };
      map.set(key, item);
    }
    const locKey =
      (location.address || '') + '|' + (location.regionKey || '');
    if (!item.locations.some((l) => (l.address || '') + '|' + (l.regionKey || '') === locKey)) {
      const loc = {};
      if (location.address != null && location.address !== '') {
        loc.address = String(location.address);
      }
      if (location.regionKey != null && location.regionKey !== '') {
        loc.regionKey = String(location.regionKey);
      }
      if (Object.keys(loc).length) item.locations.push(loc);
    }
    return item;
  }

  // ─── UI ──────────────────────────────────────────────────────────────────

  function log(msg, kind) {
    const el = document.getElementById('erz-log');
    if (!el) {
      console.log('[ERZ]', msg);
      return;
    }
    const line = document.createElement('div');
    line.className = 'erz-log-line erz-log-' + (kind || 'info');
    const time = new Date().toLocaleTimeString('ru-RU');
    line.textContent = '[' + time + '] ' + msg;
    el.appendChild(line);
    el.scrollTop = el.scrollHeight;
  }

  function setStats() {
    const s = document.getElementById('erz-stats');
    if (!s) return;
    let gc = 0;
    let bc = 0;
    state.groups.forEach((g) => {
      gc += (g.groupCompanies && g.groupCompanies.length) || 0;
      bc += (g.brandCompanies && g.brandCompanies.length) || 0;
    });
    s.textContent =
      'Регионов выбрано: ' +
      state.regionsSelected.size +
      ' / ' +
      state.regionsRaw.length +
      ' · Групп: ' +
      state.groups.size +
      ' (выбрано ' +
      state.groupsSelected.size +
      ')' +
      ' · Компаний группы: ' +
      gc +
      ' · Бренд: ' +
      bc;
  }

  function setBusy(busy) {
    state.busy = busy;
    document.querySelectorAll('#' + ROOT_ID + ' button[data-erz-action]').forEach((btn) => {
      if (btn.getAttribute('data-erz-action') === 'abort') {
        btn.disabled = !busy;
      } else if (btn.getAttribute('data-erz-action') === 'close') {
        btn.disabled = false;
      } else {
        btn.disabled = busy;
      }
    });
  }

  function buildModal() {
    const old = document.getElementById(ROOT_ID);
    if (old) old.remove();

    const root = document.createElement('div');
    root.id = ROOT_ID;
    root.innerHTML = `
<style>
#${ROOT_ID}{all:initial;font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif}
#${ROOT_ID} *{box-sizing:border-box}
#${ROOT_ID} .erz-backdrop{position:fixed;inset:0;background:rgba(15,23,42,.55);z-index:2147483000;display:flex;align-items:center;justify-content:center;padding:16px}
#${ROOT_ID} .erz-modal{width:min(920px,100%);max-height:92vh;overflow:auto;background:#f8fafc;color:#0f172a;border-radius:12px;box-shadow:0 20px 50px rgba(0,0,0,.35);border:1px solid #cbd5e1}
#${ROOT_ID} .erz-head{display:flex;justify-content:space-between;align-items:center;padding:14px 18px;background:#0f172a;color:#f8fafc;border-radius:12px 12px 0 0}
#${ROOT_ID} .erz-head h1{margin:0;font-size:16px;font-weight:700;letter-spacing:.02em}
#${ROOT_ID} .erz-body{padding:14px 18px 18px;display:flex;flex-direction:column;gap:12px}
#${ROOT_ID} .erz-row{display:flex;flex-wrap:wrap;gap:10px;align-items:center}
#${ROOT_ID} label{font-size:13px;color:#334155}
#${ROOT_ID} input[type=number]{width:80px;padding:6px 8px;border:1px solid #94a3b8;border-radius:6px;background:#fff}
#${ROOT_ID} select{padding:6px 8px;border:1px solid #94a3b8;border-radius:6px;background:#fff}
#${ROOT_ID} button{padding:7px 12px;border-radius:7px;border:1px solid #64748b;background:#e2e8f0;color:#0f172a;cursor:pointer;font-size:13px;font-weight:600}
#${ROOT_ID} button:hover{background:#cbd5e1}
#${ROOT_ID} button.primary{background:#0369a1;border-color:#0369a1;color:#fff}
#${ROOT_ID} button.primary:hover{background:#0284c7}
#${ROOT_ID} button.auto{background:#15803d;border-color:#15803d;color:#fff}
#${ROOT_ID} button.auto:hover{background:#16a34a}
#${ROOT_ID} button.danger{background:#9f1239;border-color:#9f1239;color:#fff}
#${ROOT_ID} button:disabled{opacity:.45;cursor:not-allowed}
#${ROOT_ID} .erz-panel{border:1px solid #cbd5e1;border-radius:8px;background:#fff;padding:10px}
#${ROOT_ID} .erz-panel h2{margin:0 0 8px;font-size:14px}
#${ROOT_ID} .erz-list{max-height:180px;overflow:auto;border:1px solid #e2e8f0;border-radius:6px;padding:6px;background:#f1f5f9}
#${ROOT_ID} .erz-list label{display:flex;gap:8px;align-items:flex-start;padding:3px 2px;font-size:12px}
#${ROOT_ID} .erz-stats{font-size:12px;color:#475569;padding:6px 8px;background:#e2e8f0;border-radius:6px}
#${ROOT_ID} #erz-log{max-height:160px;overflow:auto;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:11px;background:#0f172a;color:#e2e8f0;border-radius:8px;padding:8px}
#${ROOT_ID} .erz-log-line{margin:2px 0;white-space:pre-wrap;word-break:break-all}
#${ROOT_ID} .erz-log-req{color:#7dd3fc}
#${ROOT_ID} .erz-log-ok{color:#86efac}
#${ROOT_ID} .erz-log-err{color:#fca5a5}
#${ROOT_ID} .erz-hint{font-size:12px;color:#64748b}
</style>
<div class="erz-backdrop">
  <div class="erz-modal" role="dialog" aria-label="ERZ scraper">
    <div class="erz-head">
      <h1>ERZ scraper · erzrf.ru</h1>
      <div class="erz-row">
        <button type="button" data-erz-action="abort" class="danger" disabled>Стоп</button>
        <button type="button" data-erz-action="close">Закрыть</button>
      </div>
    </div>
    <div class="erz-body">
      <div class="erz-panel">
        <h2>Настройки</h2>
        <div class="erz-row">
          <label>Пауза, мс <input type="number" id="erz-pause" min="100" max="2000" step="50" value="400"></label>
          <label><input type="checkbox" id="erz-save-mid"> Сохранять промежуточные JSON</label>
          <label>names:
            <select id="erz-names-mode">
              <option value="one" selected>один регион на группу</option>
              <option value="all">все регионы группы</option>
            </select>
          </label>
          <label>join:
            <select id="erz-join-mode">
              <option value="all" selected>все регионы группы</option>
              <option value="one">один регион на группу</option>
            </select>
          </label>
        </div>
        <div class="erz-row" style="margin-top:8px">
          <button type="button" class="auto" data-erz-action="auto-all">Автоматически скачать всё</button>
        </div>
        <div class="erz-hint">Авто: все регионы → все группы → компании → ERZ_Full (пауза, промежуточные, names/join). Либо этапы вручную.</div>
        <div class="erz-hint">join=все: developer/join по каждому региону группы (иначе API вернёт только компании первого региона). names — аналогично.</div>
        <div class="erz-hint">При HTTP/JSON ошибках — до 3 попыток; для 504/502 пауза перед повтором увеличена (~8–16 с). Стоп прерывает и повторы.</div>
        <div class="erz-hint">costType=1 · куки текущей сессии · итог всегда: ERZ_Full_&lt;timestamp&gt;.json</div>
      </div>

      <div class="erz-panel">
        <h2>1. Регионы</h2>
        <div class="erz-row">
          <button type="button" class="primary" data-erz-action="load-regions">Загрузить регионы</button>
          <button type="button" data-erz-action="regions-all">Выбрать все</button>
          <button type="button" data-erz-action="regions-none">Снять все</button>
        </div>
        <div class="erz-list" id="erz-regions-list"><div class="erz-hint">Пока пусто — нажмите «Загрузить регионы»</div></div>
      </div>

      <div class="erz-panel">
        <h2>2. Группы (brand_count + brand/join)</h2>
        <div class="erz-row">
          <button type="button" class="primary" data-erz-action="load-groups">Запросить группы по выбранным регионам</button>
          <button type="button" data-erz-action="groups-all">Выбрать все группы</button>
          <button type="button" data-erz-action="groups-none">Снять все группы</button>
        </div>
        <div class="erz-list" id="erz-groups-list"><div class="erz-hint">Сначала выберите регионы и запросите группы</div></div>
      </div>

      <div class="erz-panel">
        <h2>3. Компании группы + бренд</h2>
        <div class="erz-row">
          <button type="button" class="primary" data-erz-action="load-companies">Запросить компании выбранных групп</button>
          <button type="button" data-erz-action="download-final">Скачать итог ERZ_Full</button>
        </div>
      </div>

      <div class="erz-stats" id="erz-stats">Регионов выбрано: 0 · Групп: 0 · Компаний группы: 0 · Бренд: 0</div>
      <div id="erz-log"></div>
    </div>
  </div>
</div>`;
    document.documentElement.appendChild(root);

    root.querySelector('[data-erz-action="close"]').addEventListener('click', () => root.remove());
    root.querySelector('[data-erz-action="abort"]').addEventListener('click', () => {
      state.abort = true;
      log('Запрошена остановка…', 'err');
    });
    root.querySelector('#erz-pause').addEventListener('change', (e) => {
      state.pauseMs = Number(e.target.value) || 400;
    });
    root.querySelector('#erz-save-mid').addEventListener('change', (e) => {
      state.saveIntermediate = !!e.target.checked;
    });
    root.querySelector('#erz-names-mode').addEventListener('change', (e) => {
      state.namesMode = e.target.value === 'all' ? 'all' : 'one';
    });
    root.querySelector('#erz-join-mode').addEventListener('change', (e) => {
      state.joinMode = e.target.value === 'one' ? 'one' : 'all';
    });

    root.querySelector('[data-erz-action="load-regions"]').addEventListener('click', () => run(loadRegions));
    root.querySelector('[data-erz-action="regions-all"]').addEventListener('click', () => selectAllRegions(true));
    root.querySelector('[data-erz-action="regions-none"]').addEventListener('click', () => selectAllRegions(false));
    root.querySelector('[data-erz-action="load-groups"]').addEventListener('click', () => run(loadGroups));
    root.querySelector('[data-erz-action="groups-all"]').addEventListener('click', () => selectAllGroups(true));
    root.querySelector('[data-erz-action="groups-none"]').addEventListener('click', () => selectAllGroups(false));
    root.querySelector('[data-erz-action="load-companies"]').addEventListener('click', () => run(loadCompanies));
    root.querySelector('[data-erz-action="auto-all"]').addEventListener('click', () => run(runAutoAll));
    root.querySelector('[data-erz-action="download-final"]').addEventListener('click', () => {
      try {
        syncSettingsFromUi();
        downloadFinal();
      } catch (e) {
        log(String(e && e.message ? e.message : e), 'err');
      }
    });

    log(
      'Модалка готова. Ручные этапы или «Автоматически скачать всё» (все регионы и группы).',
      'ok'
    );
    setStats();
  }

  function syncSettingsFromUi() {
    const pauseEl = document.getElementById('erz-pause');
    const midEl = document.getElementById('erz-save-mid');
    const namesEl = document.getElementById('erz-names-mode');
    const joinEl = document.getElementById('erz-join-mode');
    if (pauseEl) {
      const n = Number(pauseEl.value);
      state.pauseMs = Math.max(100, Math.min(2000, Number.isFinite(n) ? n : 400));
    }
    if (midEl) state.saveIntermediate = !!midEl.checked;
    if (namesEl) state.namesMode = namesEl.value === 'all' ? 'all' : 'one';
    if (joinEl) state.joinMode = joinEl.value === 'one' ? 'one' : 'all';
  }

  /**
   * Полный проход: все регионы → все группы → компании → ERZ_Full.
   * Учитывает паузу, промежуточные JSON и режим names.
   */
  async function runAutoAll() {
    syncSettingsFromUi();
    log(
      'Авто: старт (пауза=' +
        state.pauseMs +
        ' мс, промежуточные=' +
        (state.saveIntermediate ? 'да' : 'нет') +
        ', names=' +
        state.namesMode +
        ', join=' +
        state.joinMode +
        ')',
      'ok'
    );

    await loadRegions();
    checkAbort();
    selectAllRegions(true);
    log('Авто: выбраны все регионы (' + state.regionsSelected.size + ')', 'ok');
    if (!state.regionsSelected.size) {
      throw new Error('Нет активных регионов для автозагрузки');
    }

    await loadGroups();
    checkAbort();
    selectAllGroups(true);
    log('Авто: выбраны все группы (' + state.groupsSelected.size + ')', 'ok');
    if (!state.groupsSelected.size) {
      throw new Error('Не найдено групп для автозагрузки');
    }

    await loadCompanies();
    log('Авто: готово — итоговый ERZ_Full скачан.', 'ok');
  }

  async function run(fn) {
    if (state.busy) return;
    state.abort = false;
    syncSettingsFromUi();
    setBusy(true);
    try {
      await fn();
    } catch (e) {
      log(String(e && e.message ? e.message : e), 'err');
      console.error(e);
    } finally {
      setBusy(false);
      setStats();
    }
  }

  function checkAbort() {
    if (state.abort) throw new Error('Остановлено пользователем');
  }

  // ─── этап 1: регионы ─────────────────────────────────────────────────────

  async function loadRegions() {
    const data = await fetchJson(API + '/global/dictionary?type=buildings_regions');
    maybeSaveIntermediate('regions', 'all', 'dictionary', data);
    const list = Array.isArray(data) ? data : [];
    state.regionsRaw = list
      .filter((r) => r && !r.disabled && r.id && String(r.id) !== '0')
      .map((r) => ({
        id: String(r.id),
        text: String(r.text || ''),
        additional: String(r.additional || ''),
        brandCount: null,
      }));
    renderRegions();
    log('Регионов (активных): ' + state.regionsRaw.length, 'ok');
    setStats();
  }

  function renderRegions() {
    const box = document.getElementById('erz-regions-list');
    if (!box) return;
    box.innerHTML = '';
    state.regionsRaw.forEach((r) => {
      const lab = document.createElement('label');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = r.id;
      cb.checked = state.regionsSelected.has(r.id);
      cb.addEventListener('change', () => {
        if (cb.checked) state.regionsSelected.set(r.id, r);
        else state.regionsSelected.delete(r.id);
        setStats();
      });
      lab.appendChild(cb);
      const span = document.createElement('span');
      span.textContent = r.text + ' · ' + r.additional + ' · id=' + r.id;
      lab.appendChild(span);
      box.appendChild(lab);
    });
  }

  function selectAllRegions(on) {
    state.regionsSelected.clear();
    if (on) {
      state.regionsRaw.forEach((r) => state.regionsSelected.set(r.id, r));
    }
    renderRegions();
    setStats();
  }

  // ─── этап 2: группы ──────────────────────────────────────────────────────

  function upsertGroup(raw, regionText, regionKey, regionSlug) {
    const id = String(raw.id || '');
    if (!id) return;
    let g = state.groups.get(id);
    if (!g) {
      g = {
        id,
        name: String(raw.name || ''),
        urlId: String(raw.urlId || ''),
        regions: [],
        groupCompanies: [],
        brandCompanies: [],
        _gcMap: new Map(),
        _bcMap: new Map(),
      };
      state.groups.set(id, g);
    }
    const rk = String(regionKey);
    if (!g.regions.some((x) => x.regionKey === rk)) {
      g.regions.push({
        region: String(regionText || raw.region || ''),
        regionKey: rk,
        additional: String(regionSlug || ''),
      });
    }
  }

  function renderGroups() {
    const box = document.getElementById('erz-groups-list');
    if (!box) return;
    box.innerHTML = '';
    const arr = Array.from(state.groups.values()).sort((a, b) =>
      a.name.localeCompare(b.name, 'ru')
    );
    arr.forEach((g) => {
      const lab = document.createElement('label');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = g.id;
      cb.checked = state.groupsSelected.has(g.id);
      cb.addEventListener('change', () => {
        if (cb.checked) state.groupsSelected.add(g.id);
        else state.groupsSelected.delete(g.id);
        setStats();
      });
      lab.appendChild(cb);
      const span = document.createElement('span');
      span.textContent =
        g.name +
        ' · id=' +
        g.id +
        ' · регионов: ' +
        g.regions.length +
        (g.urlId ? ' · ' + g.urlId : '');
      lab.appendChild(span);
      box.appendChild(lab);
    });
  }

  function selectAllGroups(on) {
    state.groupsSelected.clear();
    if (on) {
      state.groups.forEach((_, id) => state.groupsSelected.add(id));
    }
    renderGroups();
    setStats();
  }

  async function loadGroups() {
    if (!state.regionsSelected.size) {
      throw new Error('Выберите хотя бы один регион');
    }
    state.groups.clear();
    state.groupsSelected.clear();

    const regions = Array.from(state.regionsSelected.values());
    let i = 0;
    for (const reg of regions) {
      checkAbort();
      i += 1;
      log(
        'Регион ' + i + '/' + regions.length + ': ' + reg.text + ' (' + reg.additional + ')',
        'info'
      );

      const q =
        'region=' +
        encodeURIComponent(reg.additional) +
        '&regionKey=' +
        encodeURIComponent(reg.id) +
        '&costType=' +
        COST_TYPE;

      const countData = await fetchJson(API + '/brand_count?' + q);
      await pause();
      checkAbort();
      const count = Number(countData && countData.count != null ? countData.count : 0);
      reg.brandCount = count;
      maybeSaveIntermediate('brand_count', reg.id, reg.additional, countData);
      log('  brand_count=' + count, 'ok');

      if (!count || count < 1) {
        continue;
      }

      const joinPath =
        API +
        '/brand/join?' +
        q +
        '&min=1&max=' +
        encodeURIComponent(String(count));
      const joinData = await fetchJson(joinPath);
      await pause();
      checkAbort();
      maybeSaveIntermediate('brand_join', reg.id, reg.additional, joinData);

      const list = joinData && Array.isArray(joinData.list) ? joinData.list : [];
      list.forEach((item) => {
        upsertGroup(item, item.region || reg.text, reg.id, reg.additional);
      });
      log('  групп в ответе: ' + list.length + '; всего уникальных: ' + state.groups.size, 'ok');
      setStats();
    }

    renderGroups();
    // по умолчанию не выбираем все — пользователь отмечает сам
    log('Готово. Уникальных групп: ' + state.groups.size + '. Отметьте нужные и запросите компании.', 'ok');
  }

  // ─── этап 3: компании ────────────────────────────────────────────────────

  function regionSlugFor(groupRegion) {
    // additional slug из сохранённого региона или lookup
    if (groupRegion.additional) return groupRegion.additional;
    const found = state.regionsRaw.find((r) => r.id === groupRegion.regionKey);
    return found ? found.additional : '';
  }

  async function fetchGroupCompanies(group, regionEntry) {
    const slug = regionSlugFor(regionEntry);
    if (!slug) {
      log('Нет additional (slug) для regionKey=' + regionEntry.regionKey + ', пропуск join', 'err');
      return;
    }
    const q =
      'region=' +
      encodeURIComponent(slug) +
      '&regionKey=' +
      encodeURIComponent(regionEntry.regionKey) +
      '&costType=' +
      COST_TYPE;
    const path = API + '/developer/join/' + encodeURIComponent(group.id) + '?' + q;
    const data = await fetchJson(path);
    maybeSaveIntermediate('developer_join', group.id, group.urlId || slug, data);
    const list = Array.isArray(data) ? data : data && Array.isArray(data.list) ? data.list : [];
    list.forEach((c) => {
      upsertCompany(group._gcMap, c, {
        address: c.address,
        regionKey: c.regionKey || regionEntry.regionKey,
      });
    });
    group.groupCompanies = Array.from(group._gcMap.values());
    return list.length;
  }

  async function fetchBrandCompanies(group, regionEntry) {
    const slug = regionSlugFor(regionEntry);
    if (!slug) {
      log('Нет additional для names, regionKey=' + regionEntry.regionKey, 'err');
      return 0;
    }
    const q =
      'region=' +
      encodeURIComponent(slug) +
      '&regionKey=' +
      encodeURIComponent(regionEntry.regionKey) +
      '&costType=' +
      COST_TYPE +
      '&organizationId=' +
      encodeURIComponent(group.id);
    const path = API + '/developer/names?' + q;
    const data = await fetchJson(path);
    maybeSaveIntermediate('developer_names', group.id, group.urlId || slug, data);
    const list = Array.isArray(data) ? data : data && Array.isArray(data.list) ? data.list : [];
    list.forEach((c) => {
      upsertCompany(group._bcMap, c, {
        address: c.address,
        regionKey: c.regionKey || regionEntry.regionKey,
      });
    });
    group.brandCompanies = Array.from(group._bcMap.values());
    return list.length;
  }

  async function loadCompanies() {
    if (!state.groupsSelected.size) {
      throw new Error('Выберите хотя бы одну группу');
    }
    const ids = Array.from(state.groupsSelected);
    let gi = 0;
    for (const id of ids) {
      checkAbort();
      gi += 1;
      const group = state.groups.get(id);
      if (!group || !group.regions.length) {
        log('Группа ' + id + ' без регионов — пропуск', 'err');
        continue;
      }
      // Повторный прогон по группе — с нуля
      group._gcMap = new Map();
      group._bcMap = new Map();
      group.groupCompanies = [];
      group.brandCompanies = [];

      log(
        'Группа ' + gi + '/' + ids.length + ': ' + group.name + ' (id=' + group.id + ')',
        'info'
      );

      // developer/join: API фильтрует по region — при «all» обходим все регионы группы
      const joinRegs =
        state.joinMode === 'all' ? group.regions : [group.regions[0]];
      let ji = 0;
      for (const reg of joinRegs) {
        checkAbort();
        ji += 1;
        const nJoin = await fetchGroupCompanies(group, reg);
        await pause();
        log(
          '  join ' +
            ji +
            '/' +
            joinRegs.length +
            ' (+' +
            (nJoin || 0) +
            ') → компаний группы ' +
            group.groupCompanies.length,
          'ok'
        );
      }

      // developer/names — one / all
      const namesRegs =
        state.namesMode === 'all' ? group.regions : [group.regions[0]];
      let ni = 0;
      for (const reg of namesRegs) {
        checkAbort();
        ni += 1;
        const n = await fetchBrandCompanies(group, reg);
        await pause();
        log(
          '  names ' +
            ni +
            '/' +
            namesRegs.length +
            ' (+' +
            (n || 0) +
            ') → бренд ' +
            group.brandCompanies.length,
          'ok'
        );
      }
      setStats();
    }

    log('Компании собраны. Скачиваю ERZ_Full…', 'ok');
    downloadFinal();
  }

  function serializeGroup(g) {
    return {
      id: g.id,
      name: g.name,
      urlId: g.urlId,
      regions: g.regions.map((r) => ({
        region: r.region,
        regionKey: r.regionKey,
      })),
      groupCompanies: (g.groupCompanies || []).map((c) => ({
        id: c.id,
        inn: c.inn,
        name: c.name,
        ogrn: c.ogrn,
        urlId: c.urlId,
        locations: c.locations || [],
      })),
      brandCompanies: (g.brandCompanies || []).map((c) => ({
        id: c.id,
        inn: c.inn,
        name: c.name,
        ogrn: c.ogrn,
        urlId: c.urlId,
        locations: (c.locations || []).map((l) => {
          const o = {};
          if (l.address) o.address = l.address;
          return o;
        }),
      })),
    };
  }

  function buildFinalObject() {
    const regions = Array.from(state.regionsSelected.values()).map((r) => ({
      id: r.id,
      text: r.text,
      additional: r.additional,
      brandCount: r.brandCount,
    }));

    // Если есть выбор групп — только они; иначе все собранные
    const source =
      state.groupsSelected.size > 0
        ? Array.from(state.groupsSelected)
            .map((id) => state.groups.get(id))
            .filter(Boolean)
        : Array.from(state.groups.values());

    return {
      meta: {
        exportedAt: new Date().toISOString(),
        costType: Number(COST_TYPE),
        namesMode: state.namesMode,
        joinMode: state.joinMode,
        pauseMs: state.pauseMs,
        source: window.location.href,
      },
      regions,
      groups: source.map(serializeGroup),
    };
  }

  function downloadFinal() {
    if (!state.groups.size && !state.regionsSelected.size) {
      throw new Error('Нет данных для выгрузки');
    }
    const obj = buildFinalObject();
    const name = 'ERZ_Full_' + ts() + '.json';
    downloadJson(name, obj);
    log('Сохранён ' + name + ' (групп: ' + obj.groups.length + ')', 'ok');
  }

  // ─── старт ───────────────────────────────────────────────────────────────

  if (!/erzrf\.ru$/i.test(window.location.hostname) && !/\.erzrf\.ru$/i.test(window.location.hostname)) {
    console.warn(
      '[ERZ] Скрипт рассчитан на erzrf.ru. Текущий хост: ' + window.location.hostname
    );
  }

  buildModal();
  console.log('[ERZ] DevTools scraper загружен. Используйте модалку на странице.');
})();
