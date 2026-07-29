/**
 * ERZ DevTools scraper (erzrf.ru)
 *
 * Запуск:
 * 1. Открыть https://erzrf.ru/zastroyschiki?... в браузере
 * 2. DevTools → Console → вставить содержимое этого файла → Enter
 * 3. Модалка: ручные этапы / авто / продолжение с чекпоинта
 *
 * Куки: credentials: 'include' (сессия браузера).
 * После 3 неудачных попыток запроса — пометка ошибки в группе и переход дальше.
 * Чекпоинт: пагинация по batchSize, сохранение каждые saveEvery групп.
 */
(function () {
  'use strict';

  const COST_TYPE = '1';
  const API = '/erz-rest/api/v1';
  const ROOT_ID = 'erz-devtools-scraper-root';
  const FETCH_MAX_ATTEMPTS = 3;
  const RETRY_BASE_MS = 2000;
  const RETRY_GATEWAY_MS = 8000;
  const CHECKPOINT_SCHEMA = 'erz-checkpoint-v1';
  const MAX_LOG_LINES = 800;
  const TREND_EPSILON = 0.02;

  /** @type {Record<string, any>} */
  const state = {
    pauseMs: 400,
    saveIntermediate: false,
    namesMode: 'one',
    joinMode: 'all',
    /** both | names | join — names достаточно для списка компаний (≈ join∪все регионы). */
    companiesSource: 'names',
    usePagination: true,
    batchSize: 1000,
    saveEvery: 250,
    regionsRaw: [],
    regionsSelected: new Map(),
    groups: new Map(),
    groupsSelected: new Set(),
    /** Порядок обработки компаний (id). */
    processQueue: [],
    busy: false,
    abort: false,
    requestCount: 0,
    groupsLoadedRaw: 0,
    currentGroupId: '',
    currentGroupName: '',
    currentQueryAlias: '—',
    currentNamesUniqueCount: 0,
    lastRequestMs: 0,
    requestDurationsMs: [],
    groupDurationsMs: [],
    currentGroupDurationsMs: [],
    processedInRun: 0,
    runTotalPlanned: 0,
    lastCheckpointSummary: '',
    pendingRender: false,
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
    downloadJson(
      stage +
        '_' +
        safeFilePart(id) +
        '_' +
        safeFilePart(urlId) +
        '_' +
        ts() +
        '.json',
      data
    );
  }

  function isGatewayStatus(status) {
    return status === 408 || status === 429 || status === 502 || status === 503 || status === 504;
  }

  function isClientErrorNoRetryLong(status) {
    return status === 400 || status === 404 || status === 422;
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

  function retryWaitMs(attempt, kind) {
    const pause = Math.max(100, Math.min(2000, Number(state.pauseMs) || 400));
    if (kind === 'client') {
      return Math.round(800 * attempt + pause);
    }
    const base = kind === 'gateway' ? RETRY_GATEWAY_MS : RETRY_BASE_MS;
    return Math.round(base * Math.pow(2, attempt - 1) + pause * attempt);
  }

  /**
   * GET JSON с повторами.
   * @returns {Promise<{ok:true, data:any}|{ok:false, error:string, status:number|null}>}
   */
  async function fetchJsonResult(pathWithQuery) {
    const url = pathWithQuery.startsWith('http')
      ? pathWithQuery
      : window.location.origin + pathWithQuery;

    let lastError = null;
    let lastStatus = null;

    for (let attempt = 1; attempt <= FETCH_MAX_ATTEMPTS; attempt++) {
      checkAbort();
      if (attempt === 1) log('→ ' + url, 'req');
      else log('↻ повтор ' + attempt + '/' + FETCH_MAX_ATTEMPTS + ': ' + url, 'req');

      let kind = 'normal';
      const startedAt = performance.now();
      try {
        const res = await fetch(url, {
          method: 'GET',
          credentials: 'include',
          headers: { Accept: 'application/json, text/plain, */*' },
        });
        lastStatus = res.status;
        state.requestCount += 1;

        const contentType = (res.headers.get('content-type') || '').toLowerCase();
        const rawText = await res.text();

        if (isGatewayStatus(res.status) || looksLikeHtmlError(rawText)) {
          kind = 'gateway';
        } else if (isClientErrorNoRetryLong(res.status)) {
          kind = 'client';
        }

        if (!res.ok) {
          let detail = '';
          try {
            const j = rawText ? JSON.parse(rawText) : null;
            if (j && (j.message || j.error || j.exception)) {
              detail = ' · ' + String(j.message || j.error || j.exception);
            }
          } catch (_) {
            if (rawText && rawText.length < 200) detail = ' · ' + rawText;
          }
          throw new Error(
            'HTTP ' + res.status + ' для ' + url + detail + (kind === 'gateway' ? ' (gateway)' : '')
          );
        }

        if (looksLikeHtmlError(rawText) || contentType.indexOf('text/html') >= 0) {
          kind = 'gateway';
          throw new Error('Ответ HTML вместо JSON для ' + url);
        }

        let data;
        try {
          data = rawText ? JSON.parse(rawText) : null;
        } catch (parseErr) {
          kind = looksLikeHtmlError(rawText) ? 'gateway' : 'normal';
          throw new Error(
            'Некорректный JSON для ' +
              url +
              ': ' +
              String(parseErr && parseErr.message ? parseErr.message : parseErr)
          );
        }

        const elapsed = performance.now() - startedAt;
        state.lastRequestMs = elapsed;
        state.requestDurationsMs.push(elapsed);
        state.currentGroupDurationsMs.push(elapsed);
        requestRender();
        if (attempt > 1) log('✓ повтор успешен (попытка ' + attempt + ')', 'ok');
        return { ok: true, data: data };
      } catch (e) {
        const elapsed = performance.now() - startedAt;
        state.lastRequestMs = elapsed;
        state.requestDurationsMs.push(elapsed);
        state.currentGroupDurationsMs.push(elapsed);
        requestRender();
        lastError = e;
        const msg = String(e && e.message ? e.message : e);
        if (state.abort || /остановлено пользователем/i.test(msg)) {
          throw e instanceof Error ? e : new Error(msg);
        }
        if (
          kind === 'normal' &&
          (/failed to fetch|networkerror|load failed|timeout/i.test(msg) ||
            (e && e.name === 'TypeError'))
        ) {
          kind = 'gateway';
        }

        if (attempt >= FETCH_MAX_ATTEMPTS) break;

        const waitMs = retryWaitMs(attempt, kind);
        log(
          '⚠ ошибка попытки ' +
            attempt +
            '/' +
            FETCH_MAX_ATTEMPTS +
            ': ' +
            msg +
            ' — жду ' +
            waitMs +
            ' мс' +
            (kind === 'gateway' ? ' (увеличенная пауза)' : kind === 'client' ? ' (400/клиент)' : ''),
          'err'
        );
        await sleep(waitMs);
        checkAbort();
      }
    }

    return {
      ok: false,
      error: String(lastError && lastError.message ? lastError.message : lastError || 'ошибка'),
      status: lastStatus,
    };
  }

  /** Строгий fetch (для dictionary / brand_count) — бросает после исчерпания попыток. */
  async function fetchJson(pathWithQuery) {
    const r = await fetchJsonResult(pathWithQuery);
    if (!r.ok) throw new Error(r.error);
    return r.data;
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
    const locKey = (location.address || '') + '|' + (location.regionKey || '');
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

  function ensureGroupMaps(group) {
    if (!group._gcMap) {
      group._gcMap = new Map();
      (group.groupCompanies || []).forEach((c) => {
        upsertCompany(group._gcMap, c, (c.locations && c.locations[0]) || {});
        (c.locations || []).forEach((loc) => upsertCompany(group._gcMap, c, loc));
      });
    }
    if (!group._bcMap) {
      group._bcMap = new Map();
      (group.brandCompanies || []).forEach((c) => {
        (c.locations || [{}]).forEach((loc) => upsertCompany(group._bcMap, c, loc));
      });
    }
    if (!Array.isArray(group.errors)) group.errors = [];
    if (!group.status) group.status = 'pending';
  }

  function recordGroupError(group, stage, message, extra) {
    ensureGroupMaps(group);
    const entry = {
      stage: String(stage || ''),
      message: String(message || ''),
      at: new Date().toISOString(),
    };
    if (extra && typeof extra === 'object') {
      Object.keys(extra).forEach((k) => {
        if (extra[k] != null && extra[k] !== '') entry[k] = extra[k];
      });
    }
    group.errors.push(entry);
    group.status = 'error';
    log('✗ пропуск после ошибок: [' + stage + '] ' + message, 'err');
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
    line.textContent = '[' + new Date().toLocaleTimeString('ru-RU') + '] ' + msg;
    el.appendChild(line);
    while (el.childElementCount > MAX_LOG_LINES) {
      el.removeChild(el.firstChild);
    }
    el.scrollTop = el.scrollHeight;
  }

  function fmtMs(ms) {
    const n = Number(ms) || 0;
    if (n <= 0) return '0м 00с 000мс';
    const total = Math.round(n);
    const min = Math.floor(total / 60000);
    const sec = Math.floor((total % 60000) / 1000);
    const msec = total % 1000;
    return min + 'м ' + String(sec).padStart(2, '0') + 'с ' + String(msec).padStart(3, '0') + 'мс';
  }

  function avg(arr) {
    if (!arr || !arr.length) return 0;
    let s = 0;
    for (let i = 0; i < arr.length; i++) s += Number(arr[i]) || 0;
    return s / arr.length;
  }

  function tailAvg(arr, n) {
    if (!arr || !arr.length) return 0;
    const start = Math.max(0, arr.length - n);
    return avg(arr.slice(start));
  }

  function trendMark(sampleAvg, globalAvg) {
    if (!sampleAvg || !globalAvg) return '→';
    const delta = (sampleAvg - globalAvg) / globalAvg;
    if (delta > TREND_EPSILON) return '🔺';
    if (delta < -TREND_EPSILON) return '🟢🔻';
    return '→';
  }

  function setStats() {
    const s = document.getElementById('erz-stats');
    if (!s) return;
    let gc = 0;
    let bc = 0;
    let done = 0;
    let pending = 0;
    let err = 0;
    state.groups.forEach((g) => {
      gc += (g.groupCompanies && g.groupCompanies.length) || 0;
      bc += (g.brandCompanies && g.brandCompanies.length) || 0;
      if (g.status === 'done') done += 1;
      else if (g.status === 'error' || g.status === 'partial') err += 1;
      else pending += 1;
    });
    const reqAvgAll = avg(state.requestDurationsMs);
    const reqAvg10 = tailAvg(state.requestDurationsMs, 10);
    const reqAvg100 = tailAvg(state.requestDurationsMs, 100);
    const grpAvgAll = avg(state.groupDurationsMs);
    const grpAvgCurrent = avg(state.currentGroupDurationsMs);
    const totalSelected = state.groupsSelected.size || state.groups.size;
    s.innerHTML =
      '<b>Регионы:</b> ' +
      state.regionsSelected.size +
      '/' +
      state.regionsRaw.length +
      ' · <b>Группы:</b> загружено ' +
      state.groupsLoadedRaw +
      ' → уникальных ' +
      state.groups.size +
      ' (выбр. ' +
      totalSelected +
      '; done ' +
      done +
      ', err/partial ' +
      err +
      ', pending ' +
      pending +
      ')<br>' +
      '<b>Текущий:</b> ' +
      (state.currentGroupName || '—') +
      (state.currentGroupId ? ' (id=' + state.currentGroupId + ')' : '') +
      ' · <b>Запрос:</b> ' +
      state.currentQueryAlias +
      ' · <b>names уник. id:</b> ' +
      state.currentNamesUniqueCount +
      '<br>' +
      '<b>Прогресс:</b> ' +
      state.processedInRun +
      '/' +
      state.runTotalPlanned +
      ' обработано · <b>Компаний:</b> group=' +
      gc +
      ' brand=' +
      bc +
      '<br>' +
      '<b>Время:</b> last=' +
      fmtMs(state.lastRequestMs) +
      ' · avg group=' +
      fmtMs(grpAvgCurrent || grpAvgAll) +
      ' · avg all=' +
      fmtMs(reqAvgAll) +
      ' · avg10=' +
      fmtMs(reqAvg10) +
      ' ' +
      trendMark(reqAvg10, reqAvgAll) +
      ' · avg100=' +
      fmtMs(reqAvg100) +
      ' ' +
      trendMark(reqAvg100, reqAvgAll) +
      '<br>' +
      '<b>Запросов:</b> ' +
      state.requestCount +
      (state.lastCheckpointSummary ? ' · <b>Чекпоинт:</b> ' + state.lastCheckpointSummary : '');
  }

  function requestRender() {
    if (state.pendingRender) return;
    state.pendingRender = true;
    requestAnimationFrame(() => {
      state.pendingRender = false;
      setStats();
    });
  }

  function setBusy(busy) {
    state.busy = busy;
    document.querySelectorAll('#' + ROOT_ID + ' button[data-erz-action]').forEach((btn) => {
      const act = btn.getAttribute('data-erz-action');
      if (act === 'abort') btn.disabled = !busy;
      else if (act === 'close' || act === 'pick-checkpoint') btn.disabled = false;
      else btn.disabled = busy;
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
#${ROOT_ID} .erz-modal{width:min(960px,100%);max-height:92vh;overflow:auto;background:#f8fafc;color:#0f172a;border-radius:12px;box-shadow:0 20px 50px rgba(0,0,0,.35);border:1px solid #cbd5e1}
#${ROOT_ID} .erz-head{display:flex;justify-content:space-between;align-items:center;padding:14px 18px;background:#0f172a;color:#f8fafc;border-radius:12px 12px 0 0}
#${ROOT_ID} .erz-head h1{margin:0;font-size:16px;font-weight:700}
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
#${ROOT_ID} button.warn{background:#a16207;border-color:#a16207;color:#fff}
#${ROOT_ID} button.danger{background:#9f1239;border-color:#9f1239;color:#fff}
#${ROOT_ID} button:disabled{opacity:.45;cursor:not-allowed}
#${ROOT_ID} .erz-panel{border:1px solid #cbd5e1;border-radius:8px;background:#fff;padding:10px}
#${ROOT_ID} .erz-panel h2{margin:0 0 8px;font-size:14px}
#${ROOT_ID} .erz-list{max-height:160px;overflow:auto;border:1px solid #e2e8f0;border-radius:6px;padding:6px;background:#f1f5f9}
#${ROOT_ID} .erz-list label{display:flex;gap:8px;align-items:flex-start;padding:3px 2px;font-size:12px}
#${ROOT_ID} .erz-stats{font-size:12px;color:#475569;padding:6px 8px;background:#e2e8f0;border-radius:6px}
#${ROOT_ID} .erz-stats b{color:#0f172a}
#${ROOT_ID} #erz-log{max-height:180px;overflow:auto;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:11px;background:#0f172a;color:#e2e8f0;border-radius:8px;padding:8px}
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
          <label><input type="checkbox" id="erz-save-mid"> Промежуточные JSON API</label>
          <label>источник компаний:
            <select id="erz-companies-source" title="names ≈ полный список; join нужен для адресов по регионам">
              <option value="names" selected>только names</option>
              <option value="join">только join</option>
              <option value="both">names + join</option>
            </select>
          </label>
          <label>names:
            <select id="erz-names-mode">
              <option value="one" selected>один регион</option>
              <option value="all">все регионы</option>
            </select>
          </label>
          <label>join:
            <select id="erz-join-mode">
              <option value="all" selected>все регионы</option>
              <option value="one">один регион</option>
            </select>
          </label>
        </div>
        <div class="erz-row" style="margin-top:8px">
          <label><input type="checkbox" id="erz-use-pagination" checked> Пагинация групп</label>
          <label>Пачка групп <input type="number" id="erz-batch" min="50" max="10000" step="50" value="1000"></label>
          <label>Чекпоинт каждые N групп <input type="number" id="erz-save-every" min="1" max="5000" step="1" value="250"></label>
        </div>
        <div class="erz-row" style="margin-top:8px">
          <button type="button" class="auto" data-erz-action="auto-all">Автоматически скачать всё</button>
          <button type="button" class="warn" data-erz-action="pick-checkpoint">Загрузить чекпоинт…</button>
          <button type="button" class="warn" data-erz-action="resume-checkpoint">Продолжить с чекпоинта</button>
          <button type="button" data-erz-action="save-checkpoint">Скачать чекпоинт сейчас</button>
          <input type="file" id="erz-checkpoint-file" accept="application/json,.json" style="display:none">
        </div>
        <div class="erz-hint">После 3 ошибок запроса — пометка в группе и переход к следующей (процесс не останавливается).</div>
        <div class="erz-hint">Чекпоинт: ERZ_Checkpoint_*.json (группы + status pending/done/error) + параллельно ERZ_Full_*.json. При выключенной пагинации обрабатываются все выбранные группы за проход.</div>
        <div class="erz-hint">costType=1 · куки сессии · HTTP 400 (names) типичен при пустом join — будет помечен и пропущен.</div>
        <div class="erz-hint">Связь группа↔компании — по <b>id</b> группы, не по имени. «Только names»: список компаний ≈ union join по всем регионам; join нужен, если важны адреса locations по регионам.</div>
      </div>

      <div class="erz-panel">
        <h2>1. Регионы</h2>
        <div class="erz-row">
          <button type="button" class="primary" data-erz-action="load-regions">Загрузить регионы</button>
          <button type="button" data-erz-action="regions-all">Выбрать все</button>
          <button type="button" data-erz-action="regions-none">Снять все</button>
        </div>
        <div class="erz-list" id="erz-regions-list"><div class="erz-hint">Пока пусто</div></div>
      </div>

      <div class="erz-panel">
        <h2>2. Группы (brand_count + brand/join)</h2>
        <div class="erz-row">
          <button type="button" class="primary" data-erz-action="load-groups">Запросить группы</button>
          <button type="button" data-erz-action="groups-all">Выбрать все группы</button>
          <button type="button" data-erz-action="groups-none">Снять все группы</button>
        </div>
        <div class="erz-list" id="erz-groups-list"><div class="erz-hint">Сначала регионы</div></div>
      </div>

      <div class="erz-panel">
        <h2>3. Компании группы + бренд</h2>
        <div class="erz-row">
          <button type="button" class="primary" data-erz-action="load-companies">Запросить компании (выбранные / pending)</button>
          <button type="button" data-erz-action="download-final">Скачать итог ERZ_Full</button>
        </div>
      </div>

      <div class="erz-stats" id="erz-stats">—</div>
      <div id="erz-log"></div>
    </div>
  </div>
</div>`;
    document.documentElement.appendChild(root);

    root.querySelector('[data-erz-action="close"]').addEventListener('click', () => root.remove());
    root.querySelector('[data-erz-action="abort"]').addEventListener('click', () => {
      state.abort = true;
      log('Запрошена остановка (сохраним чекпоинт при выходе из цикла)…', 'err');
    });

    const fileInput = root.querySelector('#erz-checkpoint-file');
    root.querySelector('[data-erz-action="pick-checkpoint"]').addEventListener('click', () => {
      fileInput.click();
    });
    fileInput.addEventListener('change', async (ev) => {
      const file = ev.target.files && ev.target.files[0];
      if (!file) return;
      try {
        const text = await file.text();
        applyCheckpoint(JSON.parse(text));
        log('Чекпоинт загружен: ' + file.name, 'ok');
        setStats();
      } catch (e) {
        log('Не удалось загрузить чекпоинт: ' + (e && e.message ? e.message : e), 'err');
      }
      fileInput.value = '';
    });

    root.querySelector('[data-erz-action="resume-checkpoint"]').addEventListener('click', () =>
      run(resumeFromCheckpoint)
    );
    root.querySelector('[data-erz-action="save-checkpoint"]').addEventListener('click', () => {
      try {
        syncSettingsFromUi();
        downloadCheckpointAndPartial('manual');
      } catch (e) {
        log(String(e && e.message ? e.message : e), 'err');
      }
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

    log('Модалка готова: авто / ручные этапы / чекпоинт + продолжение.', 'ok');
    setStats();
  }

  function syncSettingsFromUi() {
    const pauseEl = document.getElementById('erz-pause');
    const midEl = document.getElementById('erz-save-mid');
    const namesEl = document.getElementById('erz-names-mode');
    const joinEl = document.getElementById('erz-join-mode');
    const srcEl = document.getElementById('erz-companies-source');
    const paginationEl = document.getElementById('erz-use-pagination');
    const batchEl = document.getElementById('erz-batch');
    const everyEl = document.getElementById('erz-save-every');
    if (pauseEl) {
      const n = Number(pauseEl.value);
      state.pauseMs = Math.max(100, Math.min(2000, Number.isFinite(n) ? n : 400));
    }
    if (midEl) state.saveIntermediate = !!midEl.checked;
    if (namesEl) state.namesMode = namesEl.value === 'all' ? 'all' : 'one';
    if (joinEl) state.joinMode = joinEl.value === 'one' ? 'one' : 'all';
    if (srcEl) {
      const v = String(srcEl.value || 'names');
      state.companiesSource = v === 'join' || v === 'both' ? v : 'names';
    }
    if (paginationEl) state.usePagination = !!paginationEl.checked;
    if (batchEl) {
      const n = Number(batchEl.value);
      state.batchSize = Math.max(50, Math.min(10000, Number.isFinite(n) ? n : 1000));
    }
    if (everyEl) {
      const n = Number(everyEl.value);
      state.saveEvery = Math.max(1, Math.min(5000, Number.isFinite(n) ? n : 250));
    }
  }

  async function run(fn) {
    if (state.busy) return;
    state.abort = false;
    syncSettingsFromUi();
    setBusy(true);
    try {
      await fn();
    } catch (e) {
      const msg = String(e && e.message ? e.message : e);
      log(msg, 'err');
      console.error(e);
      if (/остановлено пользователем/i.test(msg) && state.groups.size) {
        try {
          downloadCheckpointAndPartial('abort');
        } catch (_) {}
      }
    } finally {
      setBusy(false);
      setStats();
    }
  }

  function checkAbort() {
    if (state.abort) throw new Error('Остановлено пользователем');
  }

  // ─── чекпоинт ────────────────────────────────────────────────────────────

  function serializeGroup(g) {
    return {
      id: g.id,
      name: g.name,
      urlId: g.urlId,
      status: g.status || 'pending',
      errors: Array.isArray(g.errors) ? g.errors.slice() : [],
      regions: (g.regions || []).map((r) => ({
        region: r.region,
        regionKey: r.regionKey,
        additional: r.additional || undefined,
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

  function buildFinalObject(includeAllGroups) {
    const regions = Array.from(state.regionsSelected.values()).map((r) => ({
      id: r.id,
      text: r.text,
      additional: r.additional,
      brandCount: r.brandCount,
    }));

    let source;
    if (includeAllGroups || state.groupsSelected.size === 0) {
      source = Array.from(state.groups.values());
    } else {
      source = Array.from(state.groupsSelected)
        .map((id) => state.groups.get(id))
        .filter(Boolean);
    }

    return {
      meta: {
        exportedAt: new Date().toISOString(),
        costType: Number(COST_TYPE),
        namesMode: state.namesMode,
        joinMode: state.joinMode,
        companiesSource: state.companiesSource,
        pauseMs: state.pauseMs,
        batchSize: state.batchSize,
        saveEvery: state.saveEvery,
        usePagination: state.usePagination,
        source: window.location.href,
      },
      regions,
      groups: source.map(serializeGroup),
    };
  }

  function buildCheckpointObject(reason) {
    const done = [];
    const pending = [];
    const errored = [];
    state.groups.forEach((g) => {
      if (g.status === 'done') done.push(g.id);
      else if (g.status === 'error' || g.status === 'partial') errored.push(g.id);
      else pending.push(g.id);
    });
    const full = buildFinalObject(true);
    return {
      checkpoint: {
        schema: CHECKPOINT_SCHEMA,
        savedAt: new Date().toISOString(),
        reason: reason || 'auto',
        phase: 'companies',
        batchSize: state.batchSize,
        saveEvery: state.saveEvery,
        usePagination: state.usePagination,
        pauseMs: state.pauseMs,
        namesMode: state.namesMode,
        joinMode: state.joinMode,
        companiesSource: state.companiesSource,
        processQueue: state.processQueue.slice(),
        counts: {
          total: state.groups.size,
          done: done.length,
          pending: pending.length,
          error: errored.length,
        },
        groupsLoadedRaw: state.groupsLoadedRaw,
        processedInRun: state.processedInRun,
        runTotalPlanned: state.runTotalPlanned,
      },
      meta: full.meta,
      regions: full.regions,
      groups: full.groups,
    };
  }

  function downloadCheckpointAndPartial(reason) {
    const stamp = ts();
    const cp = buildCheckpointObject(reason);
    downloadJson('ERZ_Checkpoint_' + stamp + '.json', cp);
    downloadJson('ERZ_Full_' + stamp + '.json', {
      meta: Object.assign({}, cp.meta, { fromCheckpoint: true, reason: reason || 'auto' }),
      regions: cp.regions,
      groups: cp.groups,
    });
    state.lastCheckpointSummary =
      'raw groups ' +
      state.groupsLoadedRaw +
      ' → unique ' +
      state.groups.size +
      ', done=' +
      cp.checkpoint.counts.done +
      ', pending=' +
      cp.checkpoint.counts.pending +
      ', error=' +
      cp.checkpoint.counts.error;
    log(
      'Чекпоинт + Full сохранены (' +
        stamp +
        '): done=' +
        cp.checkpoint.counts.done +
        ' pending=' +
        cp.checkpoint.counts.pending +
        ' error=' +
        cp.checkpoint.counts.error,
      'ok'
    );
    requestRender();
  }

  function applyCheckpoint(data) {
    if (!data || typeof data !== 'object') throw new Error('Пустой чекпоинт');
    const cp = data.checkpoint || {};
    if (cp.schema && cp.schema !== CHECKPOINT_SCHEMA) {
      log('Предупреждение: schema чекпоинта ' + cp.schema + ' (ожидался ' + CHECKPOINT_SCHEMA + ')', 'err');
    }

    if (cp.pauseMs != null) state.pauseMs = Number(cp.pauseMs) || state.pauseMs;
    if (cp.namesMode) state.namesMode = cp.namesMode === 'all' ? 'all' : 'one';
    if (cp.joinMode) state.joinMode = cp.joinMode === 'one' ? 'one' : 'all';
    if (cp.companiesSource) {
      const v = String(cp.companiesSource);
      state.companiesSource = v === 'join' || v === 'both' ? v : 'names';
    }
    if (cp.batchSize != null) state.batchSize = Number(cp.batchSize) || state.batchSize;
    if (cp.saveEvery != null) state.saveEvery = Number(cp.saveEvery) || state.saveEvery;
    if (cp.usePagination != null) state.usePagination = !!cp.usePagination;
    if (cp.groupsLoadedRaw != null) state.groupsLoadedRaw = Number(cp.groupsLoadedRaw) || 0;
    if (cp.processedInRun != null) state.processedInRun = Number(cp.processedInRun) || 0;
    if (cp.runTotalPlanned != null) state.runTotalPlanned = Number(cp.runTotalPlanned) || 0;

    const pauseEl = document.getElementById('erz-pause');
    const namesEl = document.getElementById('erz-names-mode');
    const joinEl = document.getElementById('erz-join-mode');
    const srcEl = document.getElementById('erz-companies-source');
    const paginationEl = document.getElementById('erz-use-pagination');
    const batchEl = document.getElementById('erz-batch');
    const everyEl = document.getElementById('erz-save-every');
    if (pauseEl) pauseEl.value = String(state.pauseMs);
    if (namesEl) namesEl.value = state.namesMode;
    if (joinEl) joinEl.value = state.joinMode;
    if (srcEl) srcEl.value = state.companiesSource;
    if (paginationEl) paginationEl.checked = !!state.usePagination;
    if (batchEl) batchEl.value = String(state.batchSize);
    if (everyEl) everyEl.value = String(state.saveEvery);

    state.regionsRaw = [];
    state.regionsSelected.clear();
    (data.regions || []).forEach((r) => {
      const row = {
        id: String(r.id),
        text: String(r.text || ''),
        additional: String(r.additional || ''),
        brandCount: r.brandCount != null ? r.brandCount : null,
      };
      state.regionsRaw.push(row);
      state.regionsSelected.set(row.id, row);
    });

    state.groups.clear();
    state.groupsSelected.clear();
    (data.groups || []).forEach((raw) => {
      const g = {
        id: String(raw.id || ''),
        name: String(raw.name || ''),
        urlId: String(raw.urlId || ''),
        regions: (raw.regions || []).map((r) => ({
          region: String(r.region || ''),
          regionKey: String(r.regionKey || ''),
          additional: r.additional ? String(r.additional) : '',
        })),
        groupCompanies: Array.isArray(raw.groupCompanies) ? raw.groupCompanies : [],
        brandCompanies: Array.isArray(raw.brandCompanies) ? raw.brandCompanies : [],
        errors: Array.isArray(raw.errors) ? raw.errors.slice() : [],
        status: raw.status || 'pending',
        _gcMap: null,
        _bcMap: null,
      };
      if (!g.id) return;
      // Восстановить maps из уже сохранённых компаний
      g._gcMap = new Map();
      g._bcMap = new Map();
      (g.groupCompanies || []).forEach((c) => {
        (c.locations && c.locations.length ? c.locations : [{}]).forEach((loc) =>
          upsertCompany(g._gcMap, c, loc)
        );
      });
      (g.brandCompanies || []).forEach((c) => {
        (c.locations && c.locations.length ? c.locations : [{}]).forEach((loc) =>
          upsertCompany(g._bcMap, c, loc)
        );
      });
      state.groups.set(g.id, g);
      state.groupsSelected.add(g.id);
    });

    if (Array.isArray(cp.processQueue) && cp.processQueue.length) {
      state.processQueue = cp.processQueue.map(String);
    } else {
      state.processQueue = Array.from(state.groups.keys());
    }

    renderRegions();
    renderGroups();
    const c = cp.counts || {};
    log(
      'Восстановлено групп: ' +
        state.groups.size +
        ' (done≈' +
        (c.done != null ? c.done : '?') +
        ', pending≈' +
        (c.pending != null ? c.pending : '?') +
        ')',
      'ok'
    );
  }

  async function resumeFromCheckpoint() {
    if (!state.groups.size) {
      throw new Error('Сначала загрузите файл чекпоинта (кнопка «Загрузить чекпоинт»)');
    }
    log(
      'Продолжение: пагинация=' +
        (state.usePagination ? 'on' : 'off') +
        ', пачка=' +
        state.batchSize +
        ', чекпоинт каждые ' +
        state.saveEvery,
      'ok'
    );
    await processCompaniesQueue({ limit: state.batchSize, onlyPending: true });
  }

  // ─── этап 1–2 ────────────────────────────────────────────────────────────

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
    if (on) state.regionsRaw.forEach((r) => state.regionsSelected.set(r.id, r));
    renderRegions();
    setStats();
  }

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
        errors: [],
        status: 'pending',
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
        '[' +
        (g.status || 'pending') +
        '] ' +
        g.name +
        ' · id=' +
        g.id +
        ' · рег.' +
        g.regions.length;
      lab.appendChild(span);
      box.appendChild(lab);
    });
  }

  function selectAllGroups(on) {
    state.groupsSelected.clear();
    if (on) state.groups.forEach((_, id) => state.groupsSelected.add(id));
    renderGroups();
    setStats();
  }

  async function loadGroups() {
    if (!state.regionsSelected.size) throw new Error('Выберите хотя бы один регион');
    state.groups.clear();
    state.groupsSelected.clear();
    state.processQueue = [];
    state.groupsLoadedRaw = 0;
    state.currentGroupId = '';
    state.currentGroupName = '';
    state.currentQueryAlias = 'brand_count / brand_join';
    requestRender();

    const regions = Array.from(state.regionsSelected.values());
    let i = 0;
    for (const reg of regions) {
      checkAbort();
      i += 1;
      log('Регион ' + i + '/' + regions.length + ': ' + reg.text + ' (' + reg.additional + ')', 'info');

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
      if (!count || count < 1) continue;

      const joinData = await fetchJson(
        API + '/brand/join?' + q + '&min=1&max=' + encodeURIComponent(String(count))
      );
      await pause();
      checkAbort();
      maybeSaveIntermediate('brand_join', reg.id, reg.additional, joinData);
      const list = joinData && Array.isArray(joinData.list) ? joinData.list : [];
      const before = state.groups.size;
      state.groupsLoadedRaw += list.length;
      list.forEach((item) => upsertGroup(item, item.region || reg.text, reg.id, reg.additional));
      const after = state.groups.size;
      log(
        '  join групп: +' +
          list.length +
          ' (сырой суммарно ' +
          state.groupsLoadedRaw +
          ') · дедуп: ' +
          before +
          ' → ' +
          after,
        'ok'
      );
      requestRender();
    }

    state.processQueue = Array.from(state.groups.keys());
    state.groups.forEach((g) => {
      g.status = 'pending';
    });
    renderGroups();
    downloadCheckpointAndPartial('after-groups');
    log('Уникальных групп: ' + state.groups.size + '. Чекпоинт после join сохранён.', 'ok');
  }

  // ─── этап 3: компании ────────────────────────────────────────────────────

  function regionSlugFor(groupRegion) {
    if (groupRegion.additional) return groupRegion.additional;
    const found = state.regionsRaw.find((r) => r.id === groupRegion.regionKey);
    return found ? found.additional : '';
  }

  async function fetchGroupCompanies(group, regionEntry) {
    const slug = regionSlugFor(regionEntry);
    if (!slug) {
      recordGroupError(group, 'developer/join', 'Нет additional (slug)', {
        regionKey: regionEntry.regionKey,
      });
      return { count: 0, before: group.groupCompanies.length || 0, after: group.groupCompanies.length || 0, failed: true };
    }
    const q =
      'region=' +
      encodeURIComponent(slug) +
      '&regionKey=' +
      encodeURIComponent(regionEntry.regionKey) +
      '&costType=' +
      COST_TYPE +
      '&sortType=cmxrating';
    const path = API + '/developer/join/' + encodeURIComponent(group.id) + '?' + q;
    const result = await fetchJsonResult(path);
    if (!result.ok) {
      recordGroupError(group, 'developer/join', result.error, {
        region: slug,
        regionKey: regionEntry.regionKey,
        httpStatus: result.status,
      });
      return { count: 0, before: group.groupCompanies.length || 0, after: group.groupCompanies.length || 0, failed: true };
    }
    const data = result.data;
    maybeSaveIntermediate('developer_join', group.id, group.urlId || slug, data);
    const list = Array.isArray(data) ? data : data && Array.isArray(data.list) ? data.list : [];
    const before = group._gcMap.size;
    list.forEach((c) => {
      upsertCompany(group._gcMap, c, {
        address: c.address,
        regionKey: c.regionKey || regionEntry.regionKey,
      });
    });
    group.groupCompanies = Array.from(group._gcMap.values());
    const after = group._gcMap.size;
    return { count: list.length, before: before, after: after, failed: false };
  }

  async function fetchBrandCompanies(group, regionEntry) {
    const slug = regionSlugFor(regionEntry);
    if (!slug) {
      recordGroupError(group, 'developer/names', 'Нет additional (slug)', {
        regionKey: regionEntry.regionKey,
      });
      return { count: 0, before: group.brandCompanies.length || 0, after: group.brandCompanies.length || 0, failed: true };
    }
    const q =
      'region=' +
      encodeURIComponent(slug) +
      '&regionKey=' +
      encodeURIComponent(regionEntry.regionKey) +
      '&costType=' +
      COST_TYPE +
      '&organizationId=' +
      encodeURIComponent(group.id) +
      '&sortType=cmxrating';
    const path = API + '/developer/names?' + q;
    const result = await fetchJsonResult(path);
    if (!result.ok) {
      recordGroupError(group, 'developer/names', result.error, {
        region: slug,
        regionKey: regionEntry.regionKey,
        organizationId: group.id,
        httpStatus: result.status,
      });
      return { count: 0, before: group.brandCompanies.length || 0, after: group.brandCompanies.length || 0, failed: true };
    }
    const data = result.data;
    maybeSaveIntermediate('developer_names', group.id, group.urlId || slug, data);
    const list = Array.isArray(data) ? data : data && Array.isArray(data.list) ? data.list : [];
    const before = group._bcMap.size;
    list.forEach((c) => {
      upsertCompany(group._bcMap, c, {
        address: c.address,
        regionKey: c.regionKey || regionEntry.regionKey,
      });
    });
    group.brandCompanies = Array.from(group._bcMap.values());
    const after = group._bcMap.size;
    state.currentNamesUniqueCount = after;
    requestRender();
    return { count: list.length, before: before, after: after, failed: false };
  }

  /**
   * @param {{ limit?: number, onlyPending?: boolean, ids?: string[] }} opts
   */
  async function processCompaniesQueue(opts) {
    const options = opts || {};
    let ids;
    if (options.ids && options.ids.length) {
      ids = options.ids.slice();
    } else if (options.onlyPending) {
      const queue = state.processQueue.length
        ? state.processQueue
        : Array.from(state.groups.keys());
      ids = queue.filter((id) => {
        const g = state.groups.get(id);
        return g && g.status === 'pending';
      });
    } else if (state.groupsSelected.size) {
      ids = Array.from(state.groupsSelected);
    } else {
      ids = state.processQueue.length
        ? state.processQueue.slice()
        : Array.from(state.groups.keys());
    }

    if (state.usePagination && options.limit != null && options.limit > 0) {
      ids = ids.slice(0, options.limit);
    }

    if (!ids.length) {
      log('Нет групп для обработки (все done или очередь пуста).', 'ok');
      downloadFinal();
      return;
    }

    log('Обработка компаний: ' + ids.length + ' групп…', 'ok');
    downloadCheckpointAndPartial('batch-start');
    let sinceSave = 0;
    let gi = 0;
    const totalSelected = state.groupsSelected.size || state.groups.size;
    state.runTotalPlanned = ids.length;
    state.processedInRun = 0;
    state.currentQueryAlias = 'подготовка';
    requestRender();

    for (const id of ids) {
      checkAbort();
      gi += 1;
      const groupStartedAt = performance.now();
      state.currentGroupDurationsMs = [];
      const group = state.groups.get(id);
      if (!group || !group.regions.length) {
        log('Группа ' + id + ' без регионов — пропуск', 'err');
        if (group) {
          recordGroupError(group, 'meta', 'Нет регионов у группы');
          group.status = 'error';
        }
        continue;
      }
      state.currentGroupId = group.id;
      state.currentGroupName = group.name;
      state.currentNamesUniqueCount = (group.brandCompanies && group.brandCompanies.length) || 0;

      // Не затирать уже собранные данные при resume error/partial — только для pending
      if (group.status === 'pending' || !group.status) {
        group._gcMap = new Map();
        group._bcMap = new Map();
        group.groupCompanies = [];
        group.brandCompanies = [];
        group.errors = [];
      } else {
        ensureGroupMaps(group);
      }

      log(
        'Группа ' +
          gi +
          '/' +
          ids.length +
          ' (из ' +
          totalSelected +
          '): ' +
          group.name +
          ' (id=' +
          group.id +
          ', ' +
          (group.status || 'pending') +
          ')',
        'info'
      );
      requestRender();

      let hadFail = false;
      const needJoin = state.companiesSource === 'join' || state.companiesSource === 'both';
      const needNames = state.companiesSource === 'names' || state.companiesSource === 'both';

      if (needJoin) {
        const joinRegs = state.joinMode === 'all' ? group.regions : [group.regions[0]];
        let ji = 0;
        for (const reg of joinRegs) {
          checkAbort();
          ji += 1;
          const r = await fetchGroupCompanies(group, reg);
          await pause();
          if (r.failed) hadFail = true;
          state.currentQueryAlias = 'developer/join ' + ji + '/' + joinRegs.length;
          log(
            '  join ' +
              ji +
              '/' +
              joinRegs.length +
              ' (+сырой ' +
              (r.count || 0) +
              ') · дедуп ' +
              (r.before || 0) +
              ' → ' +
              (r.after || group.groupCompanies.length) +
              (r.failed ? ' [ошибка]' : ''),
            r.failed ? 'err' : 'ok'
          );
          requestRender();
        }
      } else {
        log('  join пропущен (источник=names)', 'info');
      }

      if (needNames) {
        const namesRegs = state.namesMode === 'all' ? group.regions : [group.regions[0]];
        let ni = 0;
        for (const reg of namesRegs) {
          checkAbort();
          ni += 1;
          const r = await fetchBrandCompanies(group, reg);
          await pause();
          if (r.failed) hadFail = true;
          state.currentQueryAlias = 'developer/names ' + ni + '/' + namesRegs.length;
          log(
            '  names ' +
              ni +
              '/' +
              namesRegs.length +
              ' (+сырой ' +
              (r.count || 0) +
              ') · дедуп ' +
              (r.before || 0) +
              ' → ' +
              (r.after || group.brandCompanies.length) +
              (r.failed ? ' [ошибка]' : ''),
            r.failed ? 'err' : 'ok'
          );
          requestRender();
        }
      } else {
        log('  names пропущен (источник=join)', 'info');
      }

      if (hadFail) {
        group.status = group.errors && group.errors.length ? 'error' : 'partial';
        if (!group.errors.length) group.status = 'partial';
        // если есть хоть какие-то данные — partial
        if (
          (group.groupCompanies && group.groupCompanies.length) ||
          (group.brandCompanies && group.brandCompanies.length)
        ) {
          group.status = 'partial';
        }
      } else {
        group.status = 'done';
      }

      sinceSave += 1;
      state.processedInRun = gi;
      state.groupDurationsMs.push(performance.now() - groupStartedAt);
      requestRender();
      if (gi === 1 || gi % 25 === 0 || gi === ids.length) {
        renderGroups();
      }

      if (sinceSave >= state.saveEvery) {
        downloadCheckpointAndPartial('every-' + state.saveEvery);
        sinceSave = 0;
      }
    }

    downloadCheckpointAndPartial('batch-end');
    downloadFinal();
    log('Пачка завершена. Чекпоинт и ERZ_Full сохранены.', 'ok');
  }

  async function loadCompanies() {
    const selected = state.groupsSelected.size
      ? Array.from(state.groupsSelected)
      : null;
    await processCompaniesQueue({
      onlyPending: !selected,
      ids: selected,
      limit: state.usePagination ? state.batchSize : null,
    });
  }

  async function runAutoAll() {
    syncSettingsFromUi();
    log(
      'Авто: старт (пауза=' +
        state.pauseMs +
        ', source=' +
        state.companiesSource +
        ', names=' +
        state.namesMode +
        ', join=' +
        state.joinMode +
        ', batch=' +
        state.batchSize +
        ', pagination=' +
        (state.usePagination ? 'on' : 'off') +
        ', saveEvery=' +
        state.saveEvery +
        ')',
      'ok'
    );

    await loadRegions();
    checkAbort();
    selectAllRegions(true);
    if (!state.regionsSelected.size) throw new Error('Нет активных регионов');

    await loadGroups();
    checkAbort();
    selectAllGroups(true);
    if (!state.groupsSelected.size) throw new Error('Не найдено групп');

    // Полный авто: пачками только status=pending, пока не кончатся
    let guard = 0;
    while (guard < 100000) {
      checkAbort();
      const pending = Array.from(state.groups.values()).filter(
        (g) => g.status === 'pending'
      ).length;
      if (!pending) break;
      log('Авто: осталось pending: ' + pending, 'info');
      await processCompaniesQueue({
        onlyPending: true,
        limit: state.usePagination ? state.batchSize : null,
      });
      guard += state.usePagination ? state.batchSize : pending;
    }
    downloadFinal();
    log('Авто: готово.', 'ok');
  }

  function downloadFinal() {
    if (!state.groups.size && !state.regionsSelected.size) {
      throw new Error('Нет данных для выгрузки');
    }
    const obj = buildFinalObject(true);
    const name = 'ERZ_Full_' + ts() + '.json';
    downloadJson(name, obj);
    log('Сохранён ' + name + ' (групп: ' + obj.groups.length + ')', 'ok');
  }

  // ─── старт ───────────────────────────────────────────────────────────────

  if (!/erzrf\.ru$/i.test(window.location.hostname) && !/\.erzrf\.ru$/i.test(window.location.hostname)) {
    console.warn('[ERZ] Скрипт рассчитан на erzrf.ru. Хост: ' + window.location.hostname);
  }

  buildModal();
  console.log('[ERZ] DevTools scraper загружен (чекпоинты + soft-fail).');
})();
