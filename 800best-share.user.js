// ==UserScript==
// @name         800BEST - Thống kê tổng quan & Shipper (bản chia sẻ)
// @namespace    local.800best.report.share
// @version      3.35.0
// @description  Dashboard hub 800BEST: quét giao, POD2400, backlog, TikTok Shop, tiến độ shipper, bảng xếp hạng thi đua, COD, phân tích nguyên nhân POD. v3.19: so sánh hôm qua / tuần trước lấy trực tiếp từ Kanban (cùng giờ, đối chiếu tổng Kanban trước khi dùng, đệm gọn), quét giao/chưa quét/GTC so số chốt cuối ngày, POD hiển thị 2 số thập phân, bỏ AI dự báo, xu hướng 7 ngày, mốc giờ. v3.35: đổi công thức "Tỷ lệ ODR" — bỏ hẳn cách lấy số có sẵn từ Kanban (v3.33), giờ tự tính từ cột "Thời hạn giao hàng của đối tác" (field customerAgingDate) trong bảng đơn: (đơn đã tới hạn giao mà giao thành công + đơn CHƯA tới hạn mà đã giao thành công sớm) / (tổng đơn đã tới hạn + đơn chưa tới hạn mà đã giao thành công sớm), tính cho TOÀN BỘ đơn (không riêng TikTok); đường API (Tải nhanh) đã lấy sẵn field này nên không cần sửa gì thêm. Thẻ "TikTok đến hạn ODR" (đếm số đơn, dựa cột nội bộ, riêng TikTok) giữ nguyên không đổi.
// @match        https://v9.800best.com/sea-web/*
// @run-at       document-start
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @updateURL    https://raw.githubusercontent.com/Erick1939/tool-best/main/800best-share.user.js
// @downloadURL  https://raw.githubusercontent.com/Erick1939/tool-best/main/800best-share.user.js
// @grant        none
// ==/UserScript==

(() => {
  'use strict';

  /* ------------------------------------------------------------------ *
   * State
   * ------------------------------------------------------------------ */
  const REGION = '';
  const LOGO = '';   // nếu đặt ảnh logo (data URI) thì nút mở tool thành ảnh ở góc trên bên trái; để trống (mặc định) thì dùng nút chữ "📊 Thống kê 800BEST" ở giữa phía dưới
  const RAW_FETCH = typeof window.fetch === 'function' ? window.fetch.bind(window) : null;   // hàm fetch gốc, lấy trước khi tool gắn lớp bắt dữ liệu
  window.__b8sn = (window.__b8sn || 0) + 1;   // đếm số bản script cùng chạy trong trang (phát hiện cài trùng)
  const state = {
    office: '', province: '', officeList: [], recent: [], listTried: false,
    lastRefresh: null, busy: false, view: 'dash',
    ov: null,                  // số liệu trang tổng quan, chỉ dùng để đối chiếu
    emp: null,                 // {headers, rows, at}  tab Tiến độ nhân viên
    orders: null,              // {headers, rows, wbCol, at}  tab Chi tiết phát kiện
    sort: { key: 'rateTrip', dir: -1 },
    tiktok: 'TTVN',            // mã đầu vận đơn của TikTok Shop (cách nhau dấu phẩy)
    statusMap: {},             // trạng thái -> ok | fail | unscanned | other
    includePII: false,         // luôn tắt: không giữ SĐT/địa chỉ người nhận
    hist: new Map(), histReady: false, histErr: '', pending: null, lastSaved: null,
    hv: { office: '', mode: 'trend', preset: 'd30', from: '', to: '', group: 'day', cmpDate: '', ssort: { key: 'rateTrip', dir: -1 } },
    anm: { sel: null, base: '', cmp: 'slot2', slot: '', lastAuto: 0, authErr: '', rankKey: 'pod', trendKey: 'pod', hide: new Set(), open: '', loading: null, results: [], detail: '' },
    cod: { by: {}, meta: {}, err: {}, warn: {}, running: false, at: 0, open: false, copied: 0, zero: '', zcopied: 0, ucopied: 0 },   // COD ship nộp tiền (chỉ giữ trong bộ nhớ, không lưu)
    sh: { map: new Map(), ready: false, run: { w: false, m: false }, prog: { w: '', m: '' }, res: {}, msg: '', err0: '', hide: false },   // cột tỉ lệ giao tuần/tháng của bảng Shipper
    pick: { map: new Map(), ready: false, run: false, prog: '', msg: '', err: {}, err0: '', tried: {}, date: '', cmp: 'both', off: '', shop: '', top: false, at: 0, mon: { run: false, done: 0, total: 0, office: '' } },   // tab Pickup-sale
    wx: { data: null, at: 0, err: '', loading: false, key: '' },   // thời tiết (lượng mưa theo giờ); on/lat/lon nạp từ cấu hình
    sign: null,                // chẩn đoán trường giờ ký nhận: { keys, okN, signN }
    kb: { map: new Map(), ready: false, running: false, msg: '', err: '', tried: {}, fail: {}, diag: {}, today: new Map(), todayTried: {}, todayRun: false, want: [], sg: new Map() },   // số liệu so sánh hôm qua / tuần trước từ Kanban
    snap: { list: [], idx: new Map(), ready: false, running: false, tries: {}, last: {}, err: '', timer: 0 },   // bản chụp theo mốc giờ của tab ANM
    fast: true, apiHdr: (window.__b8shdr = window.__b8shdr || {}), sites: {}, apiNote: '', ownerSiteId: null,   // tải nhanh bằng API (chỉ ANM): tiêu đề xác thực lấy từ chính trang
    th: {},                    // ngưỡng màu/cảnh báo do người dùng chỉnh (ghi đè mặc định)
    dm: null,                  // quy tắc tính ship đi làm (nạp từ cấu hình, mặc định DM_RULES_DEFAULT)
    calc: null, calcKey: '',
    endpoints: new Map(),      // chỉ lưu tên khóa API để chẩn đoán, KHÔNG lưu giá trị
    ofList: { busy: false, at: 0, note: '', warn: false }   // kết quả nút "Tải danh sách bưu cục" ở tab Tổng quan
  };


  /* ------------------------------------------------------------------ *
   * KHÓA TAB ANM (bản chia sẻ). Mã của tab ANM được mã hóa (AES-256-GCM, khóa sinh từ mật khẩu bằng PBKDF2) và đặt trên GitHub.
   * Tổng quan và Cài đặt luôn dùng được; bấm vào ANM thì hỏi mật khẩu. File cấu hình status.json do người quản lý sửa:
   *   enabled (bật/tắt), minVersion, pwVer (đổi số này = mọi người phải nhập lại), kdf (salt, iter), payloadVer, graceDays, msg.
   * Chỉ đọc 2 file công khai trên GitHub; KHÔNG gửi bất kỳ dữ liệu nào của người dùng đi đâu.
   * ------------------------------------------------------------------ */
  const LK = { user: 'Erick1939', repo: 'tool-best', branch: 'main', ver: '3.34.1', updateUrl: 'https://raw.githubusercontent.com/Erick1939/tool-best/main/800best-share.user.js', graceDefault: 3, SKEW: 600000, RECHECK: 1800000, MAXFAIL: 5 };
  const lkS = { st: 'init', msg: '', busy: false, status: null, fails: 0, wait: 0, diag: '', checkedAt: 0 };   // st: init | checking | locked | blocked | error | open
  const lkNoop = () => undefined, lkResolved = () => Promise.resolve();
  const lkB64d = s => Uint8Array.from(atob(String(s)), c => c.charCodeAt(0));
  const lkB64e = u => { let s = ''; u.forEach(b => { s += String.fromCharCode(b); }); return btoa(s); };
  const lkVerCmp = (a, b) => { const x = String(a).split('.').map(Number), y = String(b).split('.').map(Number); for (let i = 0; i < 3; i++) { const d = (x[i] || 0) - (y[i] || 0); if (d) return d < 0 ? -1 : 1; } return 0; };
  const lkUrl = n => `https://raw.githubusercontent.com/${LK.user}/${LK.repo}/${LK.branch}/${n}?t=${Math.floor(Date.now() / 60000)}`;
  function lkStore(o) { try { if (o === undefined) return JSON.parse(localStorage.getItem('b8s_lk') || 'null') || {}; localStorage.setItem('b8s_lk', JSON.stringify(o)); } catch { /* bỏ qua */ } return {}; }
  const lkPaySave = t => { try { localStorage.setItem('b8s_lkpay', t); } catch { /* bỏ qua */ } };
  const lkPayLoad = () => { try { return localStorage.getItem('b8s_lkpay') || ''; } catch { return ''; } };
  async function lkGet(name) {
    const f = typeof RAW_FETCH === 'function' ? RAW_FETCH : (typeof fetch === 'function' ? fetch.bind(window) : null);
    if (!f) throw new Error('Trình duyệt không hỗ trợ tải cấu hình');
    let r;
    try { r = await f(lkUrl(name), { cache: 'no-store', credentials: 'omit' }); }
    catch (e) { throw Object.assign(new Error('Không kết nối được tới GitHub (mạng hoặc trang đang chặn kết nối ra ngoài)'), { net: true, detail: String(e && e.message || e) }); }
    if (!r.ok) throw Object.assign(new Error(r.status === 404 ? 'Không thấy file cấu hình trên GitHub (HTTP 404)' : `GitHub trả lỗi HTTP ${r.status}`), { net: r.status >= 500, http: r.status });
    return r.text();
  }
  function lkParseStatus(txt) {
    let o; try { o = JSON.parse(txt); } catch { throw new Error('File cấu hình trên GitHub bị hỏng (không đọc được)'); }
    if (!o || typeof o !== 'object' || !o.kdf || typeof o.kdf.salt !== 'string' || !(o.kdf.iter >= 1000) || !Number.isFinite(Number(o.pwVer))) throw new Error('File cấu hình trên GitHub sai định dạng');
    return o;
  }
  // Quyết định theo cấu hình: trả {ok:true} hoặc {ok:false, why}
  function lkJudge(st) {
    if (st.enabled === false) return { ok: false, why: String(st.msg || 'Tab ANM đang tạm khóa bởi người quản lý.') };
    if (st.minVersion && lkVerCmp(LK.ver, st.minVersion) < 0) return { ok: false, why: `Bản tool này (${LK.ver}) đã cũ, cần bản ${st.minVersion} trở lên.${LK.updateUrl ? ' Hãy cập nhật tool trong Tampermonkey.' : ''}` };
    return { ok: true };
  }
  async function lkRefreshStatus() {
    const txt = await lkGet('status.json'), st = lkParseStatus(txt), now = Date.now(), sv = lkStore();
    lkS.status = st; lkS.checkedAt = now;
    lkStore(Object.assign({}, sv, { status: st, lastOk: now, maxSeen: now }));   // vừa kết nối được GitHub: lấy giờ máy hiện tại làm mốc mới
    return st;
  }
  async function lkImportKey(raw) { return crypto.subtle.importKey('raw', raw, 'AES-GCM', true, ['decrypt']); }
  async function lkDerive(pw, st) {
    const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw.normalize('NFC')), 'PBKDF2', false, ['deriveBits']);
    return new Uint8Array(await crypto.subtle.deriveBits({ name: 'PBKDF2', salt: lkB64d(st.kdf.salt), iterations: Number(st.kdf.iter), hash: 'SHA-256' }, base, 256));
  }
  // Giải mã gói mã hóa -> chuỗi mã nguồn. Ném lỗi {bad:true} nếu khóa sai.
  async function lkDecrypt(rawKey, payText) {
    let p; try { p = JSON.parse(payText); } catch { throw new Error('Gói mã hóa trên GitHub bị hỏng'); }
    if (!p || !p.iv || !p.ct) throw new Error('Gói mã hóa trên GitHub sai định dạng');
    let buf;
    try { buf = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: lkB64d(p.iv), additionalData: new TextEncoder().encode('anm-v1') }, await lkImportKey(rawKey), lkB64d(p.ct)); }
    catch { throw Object.assign(new Error('Sai mật khẩu'), { bad: true }); }
    if (typeof DecompressionStream !== 'function') throw new Error('Trình duyệt không hỗ trợ giải nén (cần Chrome/Edge bản mới)');
    const ds = new Blob([buf]).stream().pipeThrough(new DecompressionStream('gzip'));
    return new Response(ds).text();
  }
  function lkMark() { const a = $('#b8s-app'); if (a) a.classList.toggle('lk', lkS.st !== 'open'); }
  function lkActivate(src) {
    if (!/^\/\*B8ANM/.test(src)) throw new Error('Gói mã hóa không đúng loại');
    let o;
    try { o = eval('(' + src + ')')(); }   // eslint-disable-line no-eval -- chạy trong phạm vi của tool để dùng chung các hàm
    catch (e) { throw Object.assign(new Error(`Không chạy được mã ANM (${e && e.message || e}). Có thể trang chặn chạy mã động (CSP).`), { csp: /eval|unsafe|Content Security/i.test(String(e && e.message)) }); }
    lkBind(o); lkS.st = 'open'; lkS.msg = ''; lkMark();
    try { loadAi(); loadSnapLog(); loadSnaps().then(() => { renderUI(); snapBadge(); }); } catch { /* bỏ qua */ }
    return true;
  }
  function lkRelock(st, why) {
    const was = lkS.st === 'open';
    lkUnbind(); lkS.st = st; lkS.msg = why || ''; lkMark();
    if (was) { try { renderUI(); } catch { /* bỏ qua */ } }
  }
  // Vào ứng dụng: nếu đã nhập đúng trước đó và còn trong hạn thì mở ngay, rồi kiểm tra lại với GitHub ở nền
  async function lkBoot() {
    lkMark();
    const sv = lkStore(), now = Date.now();
    const graceMs = (Number(sv.status && sv.status.graceDays) || LK.graceDefault) * 86400000;
    const cached = sv.status && sv.key && sv.status.pwVer === sv.pwVer && lkPayLoad();
    const clockOk = now >= (sv.maxSeen || 0) - LK.SKEW;
    if (cached && clockOk && now - (sv.lastOk || 0) <= graceMs && lkJudge(sv.status).ok) {
      try { lkActivate(await lkDecrypt(lkB64d(sv.key), lkPayLoad())); renderUI(); }
      catch (e) { lkS.st = 'locked'; lkS.msg = e.bad ? '' : e.message; lkMark(); }
    } else {
      lkS.st = 'locked';
      if (sv.key && sv.status && sv.status.pwVer === sv.pwVer) lkS.msg = !clockOk ? 'Đồng hồ máy bị lùi so với lần dùng trước. Hãy kết nối mạng để kiểm tra lại.' : `Đã quá ${Math.round(graceMs / 86400000)} ngày chưa kết nối được GitHub để kiểm tra. Hãy kết nối mạng, tool sẽ tự mở lại.`;
    }
    lkStore(Object.assign({}, lkStore(), { maxSeen: Math.max(now, (lkStore().maxSeen || 0)) }));
    lkRecheck(); setInterval(lkRecheck, LK.RECHECK);
    const pn = $('#b8s-anm'); if (pn) pn.addEventListener('keydown', e => { if (e.key === 'Enter' && e.target && e.target.id === 'b8s-lk-pw') lkGo(); });
  }
  // Kiểm tra lại cấu hình trên GitHub: khóa lại nếu bị tắt, đổi mật khẩu, quá hạn hoặc cần cập nhật
  async function lkRecheck() {
    if (lkS.busy) return;
    if (lkS.st === 'init') return;
    let st;
    try { st = await lkRefreshStatus(); }
    catch (e) {
      const sv = lkStore(), now = Date.now(), graceMs = (Number(sv.status && sv.status.graceDays) || LK.graceDefault) * 86400000;
      if (lkS.st === 'open' && (!(now - (sv.lastOk || 0) <= graceMs) || now < (sv.maxSeen || 0) - LK.SKEW)) lkRelock('locked', 'Đã quá hạn dùng khi không kết nối được GitHub. Hãy kết nối mạng rồi nhập lại mật khẩu.');
      else if (lkS.st === 'open') lkStore(Object.assign({}, sv, { maxSeen: Math.max(now, sv.maxSeen || 0) }));
      lkS.diag = e.detail ? `${e.message}: ${e.detail}` : e.message;
      if (state.view === 'anm') renderAnm();
      return;
    }
    const sv = lkStore(), j = lkJudge(st);
    if (sv.key && Number(sv.pwVer) !== Number(st.pwVer)) { lkStore(Object.assign({}, sv, { key: '', pwVer: null })); lkRelock('locked', 'Mật khẩu đã được đổi. Hãy nhập mật khẩu mới.'); }
    else if (!j.ok) lkRelock('blocked', j.why);
    else if (lkS.st === 'blocked' || lkS.st === 'locked') {
      // đã nhập đúng từ trước và cấu hình vẫn hợp lệ (vừa hết chặn / vừa có mạng lại): mở lại không cần nhập
      if (sv.key && Number(sv.pwVer) === Number(st.pwVer) && lkPayLoad()) { try { lkActivate(await lkDecrypt(lkB64d(sv.key), lkPayLoad())); } catch { lkS.st = 'locked'; } } else if (lkS.st === 'blocked') lkS.st = 'locked';
    }
    const sv2 = lkStore();
    if (lkS.st === 'open' && st.payloadVer && st.payloadVer !== sv2.pv && sv2.key) {
      // có bản mã mới: tải về, kiểm tra giải mã được rồi lưu để dùng từ lần mở sau
      try { const t = await lkGet('anm.dat'); await lkDecrypt(lkB64d(sv2.key), t); lkPaySave(t); lkStore(Object.assign({}, lkStore(), { pv: st.payloadVer })); log('Có bản mới của tab ANM, sẽ áp dụng khi bạn tải lại trang.'); }
      catch { /* giữ bản cũ */ }
    }
    if (state.view === 'anm' || lkS.st !== 'open') { try { renderUI(); } catch { /* bỏ qua */ } }
  }
  async function lkGo() {
    if (lkS.busy) return;
    const inp = $('#b8s-lk-pw'), pw = inp ? inp.value : '';
    if (Date.now() < lkS.wait) { lkS.msg = `Nhập sai nhiều lần. Thử lại sau ${Math.ceil((lkS.wait - Date.now()) / 1000)} giây.`; renderAnm(); return; }
    if (!pw) { lkS.msg = 'Hãy nhập mật khẩu.'; renderAnm(); return; }
    lkS.busy = true; lkS.msg = ''; renderAnm();
    try {
      if (!(window.crypto && crypto.subtle)) throw new Error('Trình duyệt không hỗ trợ mã hóa (cần mở trang bằng https)');
      const st = await lkRefreshStatus(), j = lkJudge(st);
      if (!j.ok) { lkS.st = 'blocked'; lkS.msg = j.why; return; }
      const raw = await lkDerive(pw, st), sv = lkStore();
      let pay = sv.pv && sv.pv === st.payloadVer ? lkPayLoad() : '';
      let src = '';
      if (pay) { try { src = await lkDecrypt(raw, pay); } catch (e) { if (!e.bad) throw e; pay = ''; } }
      if (!src) { pay = await lkGet('anm.dat'); src = await lkDecrypt(raw, pay); }
      lkPaySave(pay);
      lkStore(Object.assign({}, lkStore(), { key: lkB64e(raw), pwVer: Number(st.pwVer), pv: st.payloadVer || '', lastOk: Date.now() }));
      lkS.fails = 0; lkActivate(src); lkS.msg = '';
      if (inp) inp.value = '';
      log('Đã mở khóa tab ANM.');
    } catch (e) {
      if (e.bad) { lkS.fails++; lkS.msg = 'Sai mật khẩu.'; if (lkS.fails >= LK.MAXFAIL) { lkS.wait = Date.now() + 30000 * (lkS.fails - LK.MAXFAIL + 1); lkS.msg += ` Đã sai ${lkS.fails} lần, tạm khóa nhập ${30 * (lkS.fails - LK.MAXFAIL + 1)} giây.`; } }
      else { lkS.msg = e.message; lkS.diag = e.detail ? `${e.message}: ${e.detail}` : e.message; }
    } finally { lkS.busy = false; try { renderUI(); } catch { /* bỏ qua */ } }
  }
  function lkDiagText() {
    let ev = 'không'; try { new Function('return 1')(); ev = 'có'; } catch { /* bị chặn */ }
    return `Chạy mã động: ${ev} · mã hóa (crypto.subtle): ${window.crypto && crypto.subtle ? 'có' : 'KHÔNG'} · giải nén: ${typeof DecompressionStream === 'function' ? 'có' : 'KHÔNG'} · kho: ${LK.user}/${LK.repo} · bản tool ${LK.ver}${lkS.diag ? ' · lỗi gần nhất: ' + lkS.diag : ''}`;
  }
  function lkRenderAnm() {
    const box = $('#b8s-anm'); if (!box) return;
    const wait = Math.max(0, Math.ceil((lkS.wait - Date.now()) / 1000));
    const blocked = lkS.st === 'blocked';
    const html = `<div class="b8s-panel" style="max-width:520px;margin:40px auto;text-align:center;padding:26px"><h3 style="margin:0 0 6px">🔒 Tab ANM đã khóa</h3>`
      + (blocked ? `<div class="al a" style="margin:10px 0">${esc(lkS.msg)}</div>` : `<div class="muted" style="font-size:14px;margin-bottom:12px">Nhập mật khẩu để mở tab ANM. Tab Tổng quan và Cài đặt vẫn dùng bình thường.</div>`
        + `<div style="display:flex;gap:8px;justify-content:center"><input type="password" id="b8s-lk-pw" autocomplete="off" placeholder="Mật khẩu" style="width:230px;padding:7px 10px" ${lkS.busy ? 'disabled' : ''}><span class="btn" data-lk="go">${lkS.busy ? '⟳ Đang kiểm tra…' : 'Mở khóa'}</span></div>`
        + (lkS.msg ? `<div class="al ${/Sai|sai|Không|không|lỗi|quá hạn|đổi/.test(lkS.msg) ? 'r' : 'a'}" style="margin:12px 0 0">${esc(lkS.msg)}</div>` : '') + (wait ? `<div class="muted">Chờ ${wait} giây…</div>` : ''))
      + `<div style="margin-top:14px"><small class="muted" data-lk="diag" style="cursor:pointer;text-decoration:underline">Chẩn đoán</small></div><div id="b8s-lk-diag" class="muted" style="font-size:11.5px;margin-top:6px;word-break:break-word"></div></div>`;
    if (box.dataset.lk !== html) { const keep = $('#b8s-lk-pw') ? $('#b8s-lk-pw').value : ''; box.dataset.lk = html; box.innerHTML = html; const n = $('#b8s-lk-pw'); if (n && keep) n.value = keep; }
  }
  function lkClick(e) {
    const el = e.target.closest('[data-lk]'); if (!el) return;
    if (el.dataset.lk === 'go') lkGo();
    else if (el.dataset.lk === 'diag') { const d = $('#b8s-lk-diag'); if (d) d.textContent = lkDiagText(); }
  }
  function lkEnter() {
    if (lkS.st === 'open') return;
    renderAnm();
    if (!lkS.busy && Date.now() - lkS.checkedAt > 60000) lkRecheck();
  }

  // ----- Tab ANM bị khóa: các hàm của ANM là stub cho tới khi nhập đúng mật khẩu (xem lock_loader) -----
  let loadSnaps = lkResolved;
  let loadSnapLog = lkNoop;
  let snapTick = lkNoop;
  let snapBadge = lkNoop;
  let loadWeather = lkNoop;
  let renderWxInfo = lkNoop;
  let onEnterAnm = lkEnter;
  let anmOffices = lkNoop;
  let anmSel = lkNoop;
  let anmData = lkNoop;
  let fitAnm = lkNoop;
  let renderAnm = lkRenderAnm;
  let onAnmClick = lkClick;
  let onAnmChange = lkNoop;
  let addAnmSheets = lkNoop;
  function lkBind(o) { loadSnaps = o.loadSnaps; loadSnapLog = o.loadSnapLog; snapTick = o.snapTick; snapBadge = o.snapBadge; loadWeather = o.loadWeather; renderWxInfo = o.renderWxInfo; onEnterAnm = o.onEnterAnm; anmOffices = o.anmOffices; anmSel = o.anmSel; anmData = o.anmData; fitAnm = o.fitAnm; renderAnm = o.renderAnm; onAnmClick = o.onAnmClick; onAnmChange = o.onAnmChange; addAnmSheets = o.addAnmSheets; }
  function lkUnbind() { loadSnaps = lkResolved; loadSnapLog = lkNoop; snapTick = lkNoop; snapBadge = lkNoop; loadWeather = lkNoop; renderWxInfo = lkNoop; onEnterAnm = lkEnter; anmOffices = lkNoop; anmSel = lkNoop; anmData = lkNoop; fitAnm = lkNoop; renderAnm = lkRenderAnm; onAnmClick = lkClick; onAnmChange = lkNoop; addAnmSheets = lkNoop; }

  /* ------------------------------------------------------------------ *
   * Hàm thuần (không đụng DOM) – có test riêng
   * ------------------------------------------------------------------ */
  // <pure>
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function parseNum(v) {
    const s = String(v ?? '').trim();
    if (!s) return NaN;
    const m = s.match(/-?\d[\d.,]*/);
    if (!m) return NaN;
    let t = m[0];
    if (/^-?\d{1,3}([.,]\d{3})+$/.test(t)) t = t.replace(/[.,]/g, '');
    else t = t.replace(',', '.');
    return parseFloat(t);
  }

  const fmt = n => (typeof n === 'number' && Number.isFinite(n)) ? n.toLocaleString('vi-VN') : String(n ?? '');
  const round1 = v => Math.round(v * 10) / 10;
  const ratio = (a, b) => (b > 0 && Number.isFinite(a)) ? a / b * 100 : null;

  function fmtVal(v, kind) {
    if (v === null || v === undefined || !Number.isFinite(v)) return '—';
    if (kind === 'pct') return round1(v).toLocaleString('vi-VN') + '%';
    if (kind === 'dec') return round1(v).toLocaleString('vi-VN');
    return Math.round(v).toLocaleString('vi-VN');
  }

  function uniqueHeaders(list, n) {
    const seen = new Map();
    return Array.from({ length: n }, (_, i) => {
      const h = String(list[i] || `Cột ${i + 1}`).normalize('NFC');
      const c = (seen.get(h) || 0) + 1;
      seen.set(h, c);
      return c > 1 ? `${h} (${c})` : h;
    });
  }

  function toCell(v) {
    if (typeof v === 'number') return v;
    const s = String(v ?? '').trim();
    if (!/^-?\d[\d.,]*$/.test(s)) return s;
    if (/^0\d/.test(s) || s.replace(/\D/g, '').length > 9) return s;
    const n = parseNum(s);
    return Number.isFinite(n) ? n : s;
  }

  const pad = n => String(n).padStart(2, '0');
  const localStamp = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
  const safeName = s => String(s || '').replace(/[\\/:*?"<>|\s]+/g, '_').slice(0, 40);

  // ---- Vận đơn / khách ----
  const WAYBILL_RE = /^(?:[A-Za-z]{2,6}[A-Za-z0-9]{6,}|\d{10,})$/;
  const isWaybill = v => { const t = String(v ?? '').trim(); return WAYBILL_RE.test(t) && /\d/.test(t); };

  function waybillCol(headers, rows) {
    const byHeader = headers.findIndex(h => /vận đơn|waybill|运单/i.test(h));
    if (byHeader >= 0) return byHeader;
    let best = -1, bestScore = 0;
    headers.forEach((_, c) => {
      const vals = rows.slice(0, 30).map(r => r[c]);
      const score = vals.filter(isWaybill).length / Math.max(1, vals.length);
      if (score > bestScore) { bestScore = score; best = c; }
    });
    return bestScore >= 0.6 ? best : -1;
  }

  const colBy = (headers, re) => headers.findIndex(h => re.test(h));
  const parseList = text => String(text || '').split(/[,;\s]+/).map(x => x.trim().toUpperCase()).filter(Boolean);
  const isTikTok = (wb, prefixes) => { const w = String(wb ?? '').trim().toUpperCase(); return !!w && prefixes.some(p => w.startsWith(p)); };

  function sheetName(name, used) {
    const base = String(name || 'Sheet').replace(/[\[\]:*?\/\\]/g, ' ').trim().slice(0, 31) || 'Sheet';
    let out = base, i = 2;
    while (used.has(out.toLowerCase())) { const suf = ` ${i++}`; out = base.slice(0, 31 - suf.length) + suf; }
    used.add(out.toLowerCase());
    return out;
  }

  // ---- Ngày / nhân viên ----
  const DATE_RE = /(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})|(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})/;
  function parseDate(s) {
    const m = String(s ?? '').match(DATE_RE);
    if (!m) return null;
    return m[1] ? { d: +m[1], m: +m[2], y: +m[3] } : { d: +m[6], m: +m[5], y: +m[4] };
  }
  const isToday = (s, now) => {
    const p = parseDate(s);
    return !!p && p.d === now.getDate() && p.m === now.getMonth() + 1 && p.y === now.getFullYear();
  };

  // Số ngày từ ngày trong ô đến hôm nay (dương = ngày trong quá khứ); null nếu không đọc được ngày
  const dayNum = p => Math.round(Date.UTC(p.y, p.m - 1, p.d) / 86400000);
  function daysBefore(s, now) {
    const p = parseDate(s);
    return p ? dayNum({ y: now.getFullYear(), m: now.getMonth() + 1, d: now.getDate() }) - dayNum(p) : null;
  }

  // Tên không dấu, bỏ mã số/ngoặc để so khớp giữa hai tab
  const normName = s => String(s ?? '').replace(/\(.*?\)/g, ' ').replace(/[\d.]+/g, ' ')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .toLowerCase().replace(/[^a-z ]/g, ' ').replace(/\s+/g, ' ').trim();

  function shipKeys(s) {
    const raw = String(s ?? '').trim();
    if (!raw) return { code: '', name: '' };
    const m = raw.match(/\d{4,}(?:\.\d+)?/);
    return { code: m ? m[0] : '', name: normName(raw) };
  }

  // ---- Phân loại trạng thái đơn ----
  // ok = giao thành công | unscanned = nhập kho chưa quét giao | other = còn lại (đã quét giao: đang giao, thất bại, hoàn…)
  function defaultCat(status) {
    const t = String(status ?? '').normalize('NFC').toLowerCase();
    if (/đã ký nhận|giao hàng thành công/.test(t)) return 'ok';
    if (/nhập kho|chưa quét|đã đến|chờ phát|chưa phát/.test(t)) return 'unscanned';
    return 'other';
  }
  function classify(status, map) {
    const s = String(status ?? '').trim();
    const v = map && map[s];
    return (v === 'ok' || v === 'unscanned' || v === 'other') ? v : defaultCat(s);
  }

  // ---- Bưu cục / tỉnh ----
  const plain = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .toLowerCase().replace(/\s+/g, ' ').trim();
  const titleCase = s => String(s ?? '').toLocaleLowerCase('vi').replace(/(^|\s)(\S)/g, (_, a, b) => a + b.toLocaleUpperCase('vi'));

  function parseProvince(text) {
    const m = String(text ?? '').match(/Khu vực trực thuộc\s*[:：]\s*([^\n]+)/i);
    return m ? m[1].split(/\s{2,}|Thuộc khu vực|所属/)[0].trim().slice(0, 40) : '';
  }
  function parseOfficeText(text) {
    const m = String(text ?? '').match(/(?<![\p{L}])(?<!(?:Mã|Tên|Loại|Cấp)\s)Bưu cục\s*[:：]\s*([^\n]+)/iu);
    return m ? m[1].split(/\s{2,}|Thuộc khu vực|Bao gồm|所属/)[0].trim().slice(0, 60) : '';
  }

  // Chọn một dòng trong danh sách thả xuống: khớp đúng tên trước, rồi khớp phần đầu, rồi khớp duy nhất
  function pickOption(texts, name) {
    const t = plain(name), p = texts.map(plain);
    if (!t) return { index: -1, ambiguous: [] };
    const exact = p.findIndex(x => x === t);
    if (exact >= 0) return { index: exact, ambiguous: [] };
    const hits = p.map((x, k) => x.includes(t) ? k : -1).filter(k => k >= 0);
    if (hits.length === 1) return { index: hits[0], ambiguous: [] };
    const strict = hits.filter(k => p[k].startsWith(t) && !/^\s*[a-z0-9]/.test(p[k].slice(t.length)));
    if (strict.length === 1) return { index: strict[0], ambiguous: [] };
    return { index: -1, ambiguous: hits.map(k => texts[k]) };
  }

  function statusCounts(rows, stCol) {
    const m = new Map();
    for (const r of rows) { const s = String(r[stCol] ?? '').trim() || '(trống)'; m.set(s, (m.get(s) || 0) + 1); }
    return [...m].sort((a, b) => b[1] - a[1]);
  }

  // Shipper đi làm hôm nay: dòng có ngày hôm nay ở cột ngày (cột thứ 4); để trắng = không đi làm
  function workersOf(emp, now) {
    if (!emp || !emp.rows.length) return null;
    const H = emp.headers.map(h => String(h ?? '').normalize('NFC')), R = emp.rows;
    const find = l => { const i = H.findIndex(h => h === l); return i >= 0 ? i : H.findIndex(h => h.includes(l)); };
    const needCol = find('Đơn cần phát'), doneCol = find('Đã quét phát');
    let nameCol = find('Nhân viên giao nhận'); if (nameCol < 0) nameCol = 0;
    const dateCnt = H.map((_, c) => R.filter(r => parseDate(r[c])).length);
    let workCol = dateCnt[3] > 0 ? 3 : dateCnt.findIndex(n => n > 0);
    let list, note = '';
    if (workCol >= 0) {
      list = R.filter(r => isToday(r[workCol], now));
      if (!list.length) {
        list = R.filter(r => String(r[workCol] ?? '').trim() !== '');
        note = `Không dòng nào có ngày hôm nay ở cột “${H[workCol]}”; tạm tính các dòng có ngày.`;
      }
    } else {
      workCol = Math.min(3, H.length - 1);
      list = R.filter(r => String(r[workCol] ?? '').trim() !== '');
      note = 'Không thấy cột ngày; tạm tính các dòng có dữ liệu ở cột thứ 4.';
    }
    return { list, nameCol, needCol, doneCol, workCol, deptCol: find('Chức vụ'), workHeader: H[workCol] || `Cột ${workCol + 1}`, note };
  }


  // ---- Giờ giao thành công: trường thời gian của đơn (tên trường chưa được xác nhận nên thử lần lượt) ----
  const SIGN_KEYS = ['signTime', 'signingTime', 'signDate', 'signedTime', 'signOffTime', 'finishTime', 'finishDeliveryTime', 'deliveryFinishTime', 'deliveredTime', 'deliverTime', 'deliveryTime', 'latestScanTime', 'lastScanTime', 'scanTime'];
  const msOf = v => { const n = (typeof v === 'string' && /^\d{10,13}$/.test(v.trim())) ? Number(v) : v; return (typeof n === 'number' && Number.isFinite(n) && n >= 1e11 && n < 4e12) ? n : null; };
  function signMsOf(it) { for (const k of SIGN_KEYS) { const m = msOf(it && it[k]); if (m !== null) return m; } return null; }
  // Tên các trường trông như mốc thời gian (chỉ tên và số đơn có giá trị, không lấy giá trị) – để chẩn đoán
  function timeKeysOf(items) {
    const cnt = new Map();
    for (const it of (items || []).slice(0, 300)) {
      if (!it || typeof it !== 'object') continue;
      for (const [k, v] of Object.entries(it)) if (msOf(v) !== null || (/time|date/i.test(k) && v !== null && v !== undefined && v !== '')) cnt.set(k, (cnt.get(k) || 0) + 1);
    }
    return [...cnt].sort((a, b) => b[1] - a[1]).slice(0, 40).map(([k, n]) => `${k}(${n})`);
  }
  // Phút trong ngày (0..1439) của chuỗi “dd/mm/yyyy hh:mm:ss”; null nếu không đọc được
  function minOfDay(s) { const m = String(s ?? '').match(/(?:^|\s|T)(\d{1,2}):(\d{2})/); return m && +m[1] < 24 && +m[2] < 60 ? +m[1] * 60 + +m[2] : null; }
  const hmOf = m => Number.isFinite(m) ? `${pad(Math.floor(m / 60))}:${pad(Math.round(m) % 60)}` : '—';

  // ---- Quy tắc tính ship đi làm ----
  // Chỉ tính ship có chức vụ trong danh sách, quét đơn đầu tiên không muộn hơn mốc cutoff; cuối ngày (từ mốc cuối cùng) loại ship giao thành công ít hơn minOk đơn
  const DM_RULES_DEFAULT = {
    on: true,
    depts: ['Nhân viên giao nhận - Chính thức', 'Nhân viên giao nhận - Chính thức (CARGO)', 'Nhân viên giao nhận - Thuê ngoài dài hạn', 'Nhân viên giao nhận - Thuê ngoài ngắn hạn'],
    cutoff: '14:00', minOk: 5
  };
  const normDept = s => plain(String(s ?? '').replace(/\uFF08/g, '(').replace(/\uFF09/g, ')')).replace(/\s*-\s*/g, ' - ').replace(/\s*\(\s*/g, ' (').replace(/\s*\)/g, ')').replace(/\s+/g, ' ').trim();
  const hmToMin = t => { const m = /^(\d{1,2}):(\d{2})$/.exec(String(t ?? '').trim()); return m && +m[1] < 24 && +m[2] < 60 ? +m[1] * 60 + +m[2] : NaN; };
  // Cuối ngày = từ mốc tự lưu cuối cùng trở đi (23:30)
  const dmFinal = now => now.getHours() * 60 + now.getMinutes() >= slotMin(SNAP_SLOTS[SNAP_SLOTS.length - 1]);
  function dmRulesOf(r) {
    const o = Object.assign({}, DM_RULES_DEFAULT, r && typeof r === 'object' ? r : {});
    o.depts = Array.isArray(o.depts) ? o.depts.map(x => String(x).trim()).filter(Boolean) : DM_RULES_DEFAULT.depts.slice();
    o.on = o.on !== false; o.minOk = Number.isFinite(+o.minOk) ? +o.minOk : 5;
    if (!Number.isFinite(hmToMin(o.cutoff))) o.cutoff = '14:00';
    return o;
  }

  // ---- Dữ liệu API -> bảng giống bảng trên trang (để dùng lại computeDashboard) ----
  const STATUS_TXT = { 2: 'Đã nhập kho nhưng chưa quét phát', 3: 'Đang phát hàng', 4: 'Giao hàng thành công', 5: 'Giao hàng thất bại', 22: 'Chờ duyệt', 23: 'Chưa in bill hoàn', 36: 'Từ chối hoàn' };
  const vnParts = ms => { const d = new Date(ms + 7 * 3600000); return { d: pad(d.getUTCDate()), m: pad(d.getUTCMonth() + 1), y: d.getUTCFullYear(), h: pad(d.getUTCHours()), n: pad(d.getUTCMinutes()), s: pad(d.getUTCSeconds()) }; };
  const msDMY = ms => { if (!Number.isFinite(+ms) || +ms < 1e11) return ''; const t = vnParts(+ms); return `${t.d}/${t.m}/${t.y}`; };
  const msDMYHMS = ms => { if (!Number.isFinite(+ms) || +ms < 1e11) return ''; const t = vnParts(+ms); return `${t.d}/${t.m}/${t.y} ${t.h}:${t.n}:${t.s}`; };
  function apiToTables(items, empList) {
    // Mỗi dòng của API là một đơn như trên trang: KHÔNG loại dòng trùng mã hay mã rỗng (loại sẽ làm tổng đơn lệch so với trang).
    // Chỉ đếm số dòng trùng mã / mã rỗng để báo trong chẩn đoán khi tổng đơn không khớp.
    const seen = new Set(), rows = [], unknown = new Set(), codeById = new Map();
    let okN = 0, signN = 0, dupN = 0, emptyN = 0;
    for (const it of (items || [])) {
      if (it && it.dispDelivererId && it.delivererCode && !codeById.has(it.dispDelivererId)) codeById.set(it.dispDelivererId, it.delivererCode);
      const wb = String(it && it.billCode || '').trim().toUpperCase();
      if (!wb) emptyN++; else if (seen.has(wb)) dupN++; else seen.add(wb);
      const code = String(it.parcelStatus ?? '');
      let st = STATUS_TXT[code];
      if (!st) { st = `Trạng thái ${code}`; unknown.add(code); }
      const signMs = code === '4' ? signMsOf(it) : null;
      if (code === '4') { okN++; if (signMs !== null) signN++; }
      rows.push([wb, msDMY(it.predictTimeByDeparture), st, String(it.latestScanMan || '').trim(), msDMY(it.forwardAgingEndDate), signMs !== null ? msDMYHMS(signMs) : '', msDMY(it.customerAgingDate)]);
    }
    const orders = { headers: ['Mã vận đơn', 'Ngày cần phát hàng', 'Trạng thái đơn hàng', 'Người quét mới nhất', 'Thời hạn giao hàng nội bộ', 'Giờ giao thành công', 'Thời hạn giao hàng của đối tác'], rows, wbCol: 0 };
    const eRows = (empList || []).map(e => {
      const c = codeById.get(e.dispDelivererId);
      return [`${String(e.delivererName || '').trim()}${c ? `(${c})` : ''}`, e.shouldDispCount ?? 0, e.todayDispCount ?? 0, msDMYHMS(e.firstDispScanTime), e.signCount ?? 0, String(e.departmentName ?? '').trim()];
    });
    const emp = { headers: ['Nhân viên giao nhận', 'Đơn cần phát', 'Đã quét phát', 'Quét phát đầu tiên', 'Đã ký nhận', 'Chức vụ'], rows: eRows };
    return { orders, emp, unknown: [...unknown], signKeys: timeKeysOf(items), okN, signN, dupN, emptyN };
  }

  // ---- Tính dashboard ----
  function computeDashboard({ orders, emp, tiktok, statusMap, now, rules, skip }) {
    const H = orders.headers.map(h => String(h ?? '').normalize('NFC')), wb = orders.wbCol, rows = orders.rows;
    const stCol = colBy(H, /trạng thái/i);
    const scCol = colBy(H, /người quét|nhân viên|shipper|người phát|giao nhận/i);
    const dCol = colBy(H, /ngày cần phát/i);
    // Cột “Thời hạn giao hàng nội bộ”: ưu tiên tìm theo tên; nếu tiêu đề không đọc được thì suy ra từ cột ngày nằm ngay trước “Quá thời hạn giao hàng nội bộ”
    const looksDate = c => { const v = rows.map(r => r[c]).filter(x => String(x ?? '').trim()).slice(0, 60); return v.length >= 3 && v.filter(x => DATE_RE.test(String(x))).length / v.length >= 0.6; };
    let oCol = H.findIndex(h => /(thời\s*(hạn|gian)|hạn).*nội\s*bộ|odr/i.test(h) && !/quá|còn lại|chênh/i.test(h));
    if (oCol < 0) {
      const q = H.findIndex(h => /quá\s*thời\s*(hạn|gian).*nội\s*bộ/i.test(h));
      if (q > 0 && looksDate(q - 1)) oCol = q - 1;
    }
    // Nếu không có cột ngày thì dùng cột “Quá thời hạn giao hàng nội bộ” = 0 (đến hạn đúng hôm nay)
    let xCol = -1;
    if (oCol < 0) xCol = H.findIndex(h => /quá\s*thời\s*(hạn|gian).*nội\s*bộ/i.test(h));
    // Cột “Thời hạn giao hàng của đối tác” (v3.35, dùng để tính Tỷ lệ ODR): ưu tiên tìm theo tên; nếu không đọc được tiêu đề thì suy ra từ cột ngày nằm ngay trước “Quá thời hạn giao hàng của đối tác”
    let pCol = H.findIndex(h => /thời\s*(hạn|gian).*(của\s*)?đối\s*tác/i.test(h) && !/quá|còn lại|chênh/i.test(h));
    if (pCol < 0) {
      const q = H.findIndex(h => /quá\s*thời\s*(hạn|gian).*đối\s*tác/i.test(h));
      if (q > 0 && looksDate(q - 1)) pCol = q - 1;
    }
    // Không có cột ngày thì dùng cột text “Quá/Còn … thời hạn giao hàng của đối tác”: “Quá” = đã tới hạn, “Còn” = chưa tới hạn
    let pxCol = -1;
    if (pCol < 0) pxCol = H.findIndex(h => /quá\s*thời\s*(hạn|gian).*đối\s*tác|thời\s*(hạn|gian).*đối\s*tác.*(quá|còn)/i.test(h));
    const tCol = colBy(H, /giờ giao thành công/i);
    const bl = { total: 0, wait: 0, fail: 0, unscanned: 0, d5: 0 };
    let odr = 0, firstOk = null;
    const odr2 = { due: 0, ok: 0 };
    const mkAgg = () => ({ total: 0, ok: 0, unscanned: 0, newTotal: 0, newOk: 0, newUns: 0 });
    const all = mkAgg(), tt = mkAgg(), sme = mkAgg();
    const byCode = new Map(), byName = new Map(), stats = [];

    const getStat = cell => {
      const k = shipKeys(cell);
      if (!k.code && !k.name) return null;
      let o = k.code ? byCode.get(k.code) : byName.get(k.name);
      if (!o) { o = { label: String(cell).trim(), total: 0, ok: 0, ttTotal: 0, ttOk: 0 }; stats.push(o); }
      if (k.code) byCode.set(k.code, o);
      if (k.name && !byName.has(k.name)) byName.set(k.name, o);
      return o;
    };

    for (const r of rows) {
      const isTT = isTikTok(r[wb], tiktok);
      const cat = stCol >= 0 ? classify(r[stCol], statusMap) : 'other';
      const dateStr = dCol >= 0 ? r[dCol] : '';
      const isNew = dCol >= 0 && isToday(dateStr, now);
      const back = dCol >= 0 ? daysBefore(dateStr, now) : null;
      for (const a of (isTT ? [all, tt] : [all, sme])) {
        a.total++;
        if (cat === 'ok') a.ok++; else if (cat === 'unscanned') a.unscanned++;
        if (isNew) { a.newTotal++; if (cat === 'ok') a.newOk++; else if (cat === 'unscanned') a.newUns++; }
      }
      // Giờ đơn giao thành công đầu tiên trong ngày (đơn thành công có giờ ký nhận là hôm nay)
      if (tCol >= 0 && cat === 'ok' && isToday(r[tCol], now)) { const mm = minOfDay(r[tCol]); if (mm !== null && (firstOk === null || mm < firstOk)) firstOk = mm; }
      // Backlog: đơn có ngày cần phát cách hôm nay từ 3 ngày trở lên và chưa giao thành công (d5: từ 5 ngày)
      if (back !== null && back >= 3 && cat !== 'ok') {
        bl.total++;
        if (back >= 5) bl.d5++;
        const st = String(stCol >= 0 ? r[stCol] : '').normalize('NFC').toLowerCase();
        if (/chờ duyệt/.test(st)) bl.wait++;
        else if (/thất bại|không thành công/.test(st)) bl.fail++;
        else if (cat === 'unscanned') bl.unscanned++;
      }
      // TikTok đến hạn ODR: thời hạn giao hàng nội bộ là hôm nay
      if (isTT && oCol >= 0 && isToday(r[oCol], now)) { odr++; }
      else if (isTT && oCol < 0 && xCol >= 0) { const m = String(r[xCol] ?? '').match(/\d+/); if (m && +m[0] === 0 && !/^-/.test(String(r[xCol]).trim())) { odr++; } }
      // v3.35: Tỷ lệ ODR = (đơn tới hạn giao hàng CỦA ĐỐI TÁC đã giao thành công + đơn CHƯA tới hạn mà đã giao thành công sớm)
      // / (tổng đơn tới hạn + đơn chưa tới hạn mà đã giao thành công sớm) — tính cho TOÀN BỘ đơn (không riêng TikTok)
      {
        let due = null;   // true = đã tới hạn (hạn <= hôm nay), false = chưa tới hạn, null = không xác định được
        if (pCol >= 0) { const db = daysBefore(r[pCol], now); if (db !== null) due = db >= 0; }
        else if (pxCol >= 0) { const s = String(r[pxCol] ?? '').trim(); if (/quá/i.test(s)) due = true; else if (/còn/i.test(s)) due = false; }
        if (due === true) { odr2.due++; if (cat === 'ok') odr2.ok++; }
        else if (due === false && cat === 'ok') { odr2.due++; odr2.ok++; }
      }
      // Đơn đã quét giao được tính cho người quét mới nhất
      if (cat !== 'unscanned' && scCol >= 0) {
        const o = getStat(r[scCol]);
        if (o) {
          o.total++;
          if (cat === 'ok') o.ok++;
          if (isTT) { o.ttTotal++; if (cat === 'ok') o.ttOk++; }
        }
      }
    }

    const W = workersOf(emp, now);
    const signIdx = W ? emp.headers.findIndex(h => h.includes('Đã ký nhận')) : -1;
    const ftIdx = W ? emp.headers.findIndex(h => h.includes('Quét phát đầu tiên')) : -1;
    const workersEstimated = !W;
    let nWork = W ? W.list.length : stats.filter(o => o.total > 0).length;   // với W: tính lại sau khi lọc ship đủ điều kiện
    const DM = dmRulesOf(rules), dmWhy = { dept: 0, late: 0, few: 0 };
    let skipped = 0;

    const used = new Set();
    let shippers;
    if (W) {
      shippers = W.list.map(row => {
        const cell = row[W.nameCol];
        const k = shipKeys(cell);
        const o = (k.code && byCode.get(k.code)) || (k.name && byName.get(k.name)) || null;
        if (o) used.add(o);
        const done = W.doneCol >= 0 ? parseNum(row[W.doneCol]) : NaN;
        return { name: cell, matched: !!o, scanned: Number.isFinite(done) ? done : (o ? o.total : 0),
          // GTC của shipper: lấy theo cột “Đã ký nhận” của trang (đúng số trang hiển thị); không có cột đó mới đếm từ bảng đơn
          ok: (signIdx >= 0 && Number.isFinite(parseNum(row[signIdx]))) ? parseNum(row[signIdx]) : (o ? o.ok : 0),
          ttTotal: o ? o.ttTotal : 0, ttOk: o ? o.ttOk : 0,
          signed: signIdx >= 0 ? parseNum(row[signIdx]) : NaN,
          ft: (ftIdx >= 0 && isToday(row[ftIdx], now)) ? minOfDay(row[ftIdx]) : null, _row: row };
      });
      // Chỉ giữ ship đủ điều kiện đi làm; ship không đủ bị ẩn (đơn của họ vẫn nằm trong các thẻ tổng)
      if (DM.on) {
        const hasDept = W.deptCol >= 0 && W.list.some(r => String(r[W.deptCol] ?? '').trim() !== '');   // bảng không có chức vụ thì không lọc theo chức vụ
        const allow = new Set(DM.depts.map(normDept)), cut = hmToMin(DM.cutoff), fin = dmFinal(now);
        shippers = shippers.filter(s => {
          if (hasDept && allow.size && !allow.has(normDept(s._row[W.deptCol]))) { dmWhy.dept++; return false; }
          if (Number.isFinite(cut) && Number.isFinite(s.ft) && s.ft > cut) { dmWhy.late++; return false; }
          if (fin && DM.minOk > 0 && !((Number.isFinite(s.ok) ? s.ok : 0) >= DM.minOk)) { dmWhy.few++; return false; }
          return true;
        });
      }
      if (typeof skip === 'function') { const n0 = shippers.length; shippers = shippers.filter(s => !skip(s.name)); skipped = n0 - shippers.length; }
      shippers.forEach(s => { delete s._row; });
      nWork = shippers.length;
    } else {
      shippers = stats.filter(o => o.total > 0).map(o => ({ name: o.label, matched: true, scanned: o.total,
        ok: o.ok, ttTotal: o.ttTotal, ttOk: o.ttOk }));
      if (typeof skip === 'function') shippers = shippers.filter(s => !skip(s.name));
    }
    shippers.forEach(s => { s.rateTrip = ratio(s.ok, s.scanned); s.rateTT = ratio(s.ttOk, s.ttTotal); });

    const sum = k => shippers.reduce((a, s) => a + (Number.isFinite(s[k]) ? s[k] : 0), 0);
    const totals = { name: 'TỔNG', scanned: sum('scanned'), ok: sum('ok'), ttTotal: sum('ttTotal'), ttOk: sum('ttOk') };
    totals.rateTrip = ratio(totals.ok, totals.scanned);
    totals.rateTT = ratio(totals.ttOk, totals.ttTotal);

    const notes = [];
    if (stCol < 0) notes.push('Không thấy cột “Trạng thái” trong bảng đơn nên chưa phân loại được thành công/chưa quét.');
    if (dCol < 0) notes.push('Không thấy cột “Ngày cần phát hàng” trong bảng đơn nên chưa tính được POD2400 và Backlog 3 ngày.');
    if (oCol < 0 && xCol < 0) notes.push('Không thấy cột “Thời hạn giao hàng nội bộ” trong bảng đơn nên chưa tính được TikTok đến hạn ODR.');
    if (pCol < 0 && pxCol < 0) notes.push('Không thấy cột “Thời hạn giao hàng của đối tác” trong bảng đơn nên chưa tính được Tỷ lệ ODR.');
    if (scCol < 0) notes.push('Không thấy cột “Người quét” trong bảng đơn nên chưa tách được theo shipper.');
    if (W && W.note) notes.push(W.note);
    if (W && DM.on) {
      const hid = W.list.length - skipped - shippers.length;
      if (!(W.deptCol >= 0 && W.list.some(r => String(r[W.deptCol] ?? '').trim() !== ''))) notes.push('Không thấy cột “Chức vụ” trong bảng nhân viên nên chưa lọc theo chức vụ.');
      if (hid > 0) notes.push(`Đã ẩn ${hid}/${W.list.length} ship không đủ điều kiện đi làm (${[dmWhy.dept ? `${dmWhy.dept} sai chức vụ` : '', dmWhy.late ? `${dmWhy.late} quét đầu sau ${DM.cutoff}` : '', dmWhy.few ? `${dmWhy.few} giao thành công dưới ${DM.minOk} đơn (cuối ngày)` : ''].filter(Boolean).join(', ')}).`);
    }
    if (!W) notes.push('Chưa có bảng Tiến độ nhân viên: số nhân viên và danh sách shipper được ước tính từ bảng đơn.');
    const unmatched = shippers.filter(s => !s.matched).length;
    if (W && unmatched) notes.push(`${unmatched} shipper đi làm không khớp được tên/mã với bảng đơn (GTC tính bằng 0).`);
    const loose = stats.filter(o => !used.has(o)).reduce((a, o) => a + o.total, 0);
    if (W && loose) notes.push(`${fmt(loose)} đơn do người không có trong danh sách đi làm quét nên không nằm trong bảng shipper (vẫn tính trong các thẻ).`);
    if (W) {
      const signCol = emp.headers.findIndex(h => h.includes('Đã ký nhận'));
      if (signCol >= 0) {
        const empSigned = W.list.reduce((a, r) => { const n = parseNum(r[signCol]); return a + (Number.isFinite(n) ? n : 0); }, 0);
        const fromOrders = all.ok, gap = Math.abs(empSigned - fromOrders);
        notes.push(gap === 0
          ? `✔ GTC từng shipper lấy theo cột “Đã ký nhận”: tổng ${fmt(empSigned)}, khớp số đơn thành công đếm từ bảng đơn.`
          : gap <= Math.max(3, Math.round(fromOrders * 0.02))
            ? `GTC từng shipper lấy theo cột “Đã ký nhận”: tổng ${fmt(empSigned)} (đếm từ bảng đơn được ${fmt(fromOrders)}, lệch ${fmt(gap)} đơn là bình thường).`
            : `⚠ Cột “Đã ký nhận” ghi tổng ${fmt(empSigned)}, đếm từ bảng đơn được ${fmt(fromOrders)} (lệch ${fmt(gap)}). Kiểm tra phân loại trạng thái ở Cài đặt.`);
      }
    }

    const card = a => ({
      pod: ratio(a.newOk, a.newTotal), newOk: a.newOk, newTotal: a.newTotal, newUns: a.newUns,
      newScan: ratio(a.newTotal - a.newUns, a.newTotal),
      total: a.total, unscanned: a.unscanned, scanned: a.total - a.unscanned, ok: a.ok,
      scanRate: ratio(a.total - a.unscanned, a.total),
      perNV: nWork > 0 ? a.total / nWork : null,
      gtcTrip: ratio(a.ok, a.total - a.unscanned),
      gtcTotal: ratio(a.ok, a.total)
    });
    return {
      cards: { all: card(all), tt: card(tt), sme: card(sme) },
      extra: { backlog: dCol >= 0 ? bl : null, odrTT: (oCol >= 0 || xCol >= 0) ? odr : null, odr2: (oCol >= 0 || xCol >= 0) ? odr2 : null },
      firstOk, hasTime: tCol >= 0,
      hasTT: tt.total > 0, nWork, workersEstimated, W, shippers, totals, notes, stCol, scCol, dmHidden: W ? W.list.length - shippers.length : 0
    };
  }

  // ---- v3.19: số liệu so sánh (hôm qua, tuần trước) tính từ danh sách đơn của API Kanban, cùng giờ chính xác đến phút ----
  const KB_NONE = -99999, KB_F = { NEW: 1, TT: 2, D24: 4 };
  const kbMs = v => { const t = Date.parse(String(v ?? '')); return Number.isFinite(t) ? t : null; };
  const kbDay0 = date => Date.UTC(+date.slice(0, 4), +date.slice(5, 7) - 1, +date.slice(8, 10)) - 7 * 3600000;   // 00:00 giờ Việt Nam của ngày
  // v3.20: giờ ký nhận (phút trong ngày) của các đơn thành công hôm nay, lấy từ bảng đơn của tool (không phụ thuộc Kanban)
  function signMinsOf(orders, statusMap, now) {
    try {
      const H = orders.headers.map(h => String(h ?? '').normalize('NFC')), tCol = colBy(H, /giờ giao thành công/i), stCol = colBy(H, /trạng thái/i);
      if (tCol < 0 || stCol < 0) return null;
      const out = [];
      for (const r of orders.rows) if (classify(r[stCol], statusMap) === 'ok' && isToday(r[tCol], now)) { const m = minOfDay(r[tCol]); if (m !== null) out.push(m); }
      return out.sort((a, b) => a - b);
    } catch { return null; }
  }
  // Nén danh sách đơn của một ngày: mỗi đơn = [phút nhập kho, phút ký nhận, cờ, số thứ tự ship] (phút tính từ 00:00 giờ VN); không giữ mã vận đơn
  function kbPack(items, tiktok, date) {
    const d0 = kbDay0(date), ships = [], idx = new Map(), P = [];
    const mn = ms => ms === null ? KB_NONE : Math.round((ms - d0) / 60000);
    for (const x of (items || [])) {
      if (!x || typeof x !== 'object') continue;
      const name = String(x.smallManName ?? '').trim();
      let si = -1;
      if (name) { si = idx.has(name) ? idx.get(name) : -1; if (si < 0) { si = ships.length; idx.set(name, si); ships.push(name); } }
      const f = (Number(x.isTodaySdisp) === 1 ? KB_F.NEW : 0) | (isTikTok(x.billCode, tiktok) ? KB_F.TT : 0) | (Number(x.is2400PodSign) === 1 ? KB_F.D24 : 0);
      P.push(mn(kbMs(x.endSiteArrTime)), x.isSign ? mn(kbMs(x.endSiteSignTime)) : KB_NONE, f, si);
    }
    return { ships, P };
  }
  // Số liệu cả ngày tính từ danh sách (để đối chiếu với số tổng của Kanban)
  function kbDayCheck(E) {
    const P = E.P; let all = 0, newN = 0, newOk = 0, ttNew = 0, ttOk = 0, ok = 0;
    for (let i = 0; i < P.length; i += 4) {
      all++; const f = P[i + 2], nw = f & KB_F.NEW;
      if (P[i + 1] !== KB_NONE && P[i + 1] < 1440) ok++;
      if (nw) { newN++; if (f & KB_F.D24) newOk++; if (f & KB_F.TT) { ttNew++; if (f & KB_F.D24) ttOk++; } }
    }
    return { all, newN, newOk, ttNew, ttOk, ok };
  }
  // Bản ghi giống bản chụp tại phút m của ngày (m >= 1440 là số cả ngày). Chỉ có số đơn/giao thành công/ship; không có số quét giao
  function kbAt(E, m, rules, office, date) {
    const P = E.P, fin = m >= 1440, lim = fin ? 1440 : m;
    const all = { total: 0, unscanned: 0, ok: 0, newTotal: 0, newOk: 0, newUns: null }, tt = { total: 0, unscanned: 0, ok: 0, newTotal: 0, newOk: 0, newUns: null };
    const sh = E.ships.map(() => ({ ok: 0, tt: 0, ft: null }));
    let first = null;
    for (let i = 0; i < P.length; i += 4) {
      const arr = P[i], sg = P[i + 1], f = P[i + 2], si = P[i + 3];
      if (!fin && arr !== KB_NONE && arr > m) continue;                  // chưa nhập kho tại thời điểm đó
      const isNew = f & KB_F.NEW, isTT = f & KB_F.TT, signed = sg !== KB_NONE && sg <= lim, okNew = fin ? !!(f & KB_F.D24) : signed;
      all.total++; if (isTT) tt.total++;
      if (signed) { all.ok++; if (isTT) tt.ok++; if (si >= 0) { sh[si].ok++; if (isTT) sh[si].tt++; if (sh[si].ft === null || sg < sh[si].ft) sh[si].ft = sg; } if (first === null || sg < first) first = sg; }
      if (isNew) { all.newTotal++; if (okNew) all.newOk++; if (isTT) { tt.newTotal++; if (okNew) tt.newOk++; } }
    }
    const DM = dmRulesOf(rules), cutDm = hmToMin(DM.cutoff), ships = [];
    const skipK = shipSkipFn(office);
    E.ships.forEach((n, i) => {
      const h = sh[i]; if (!(h.ok > 0)) return;
      if (skipK && skipK(n)) return;
      if (DM.on) {
        if (Number.isFinite(cutDm) && h.ft !== null && h.ft > cutDm) return;
        if (fin && DM.minOk > 0 && h.ok < DM.minOk) return;
      }
      ships.push({ n, sc: 0, ok: h.ok, tt: h.tt, tk: 0, sg: h.ok, ft: h.ft });
    });
    const odr2 = (E.fin && Number.isFinite(E.fin.odrDen) && E.fin.odrDen > 0) ? { due: E.fin.odrDen, ok: E.fin.odrOk || 0 } : null;
    return { office, date, id: `${office}|${date}`, savedAt: '', v: 1, verified: true, all, tt, backlog: null, odrTT: null, odr2, firstOk: first, nWork: ships.length, nScan: null, workersEstimated: true, ships, rb: 'kb', rbAt: '', fin: E.fin || null };
  }

  // Thẻ chỉ số: [nhãn, giá trị, kiểu, màu]; thẻ không có dữ liệu bị bỏ khi hiển thị
  function cardList(c) {
    const a = c.cards.all, t = c.cards.tt, s = c.cards.sme, e = c.extra || {};
    const bl = e.backlog;
    return {
      main: [
        ['Tổng đơn', a.total, 'int', 'blue'], ['Tỷ lệ quét giao', a.scanRate, 'pct', 'green'],
        ['POD2400', a.pod, 'pct', 'amber', [['Giao thành công', a.newOk], ['Tổng đơn mới hôm nay', a.newTotal]]],
        ['Chưa quét giao', a.unscanned, 'int', 'purple', [['SME', s.unscanned], ['TikTok', t.unscanned]]],
        ['GTC chuyến', a.gtcTrip, 'pct', 'teal'], ['GTC Total', a.gtcTotal, 'pct', 'orange'],
        ['Backlog 3 ngày', bl ? bl.total : null, 'int', 'red',
          bl ? [['Chờ duyệt', bl.wait], ['Giao thất bại', bl.fail], ['Nhập kho chưa quét phát', bl.unscanned]] : null]
      ],
      tts: [
        ['Tổng đơn TikTok', t.total, 'int', 'pink'], ['Tỷ lệ quét TikTok', t.scanRate, 'pct', 'green'],
        ['TikTok POD2400', t.pod, 'pct', 'amber', [['Giao thành công', t.newOk], ['Tổng đơn mới hôm nay', t.newTotal]]],
        ['TikTok chưa quét giao', t.unscanned, 'int', 'purple'],
        ['GTC chuyến TikTok', t.gtcTrip, 'pct', 'teal'], ['GTC Total TikTok', t.gtcTotal, 'pct', 'orange'],
        ['TikTok đến hạn ODR', e.odrTT, 'int', 'red']
      ]
    };
  }

  function shipperCols(c) {
    const cols = [
      { k: 'name', label: 'Mã NV - Tên NV', group: '', kind: 'text' },
      { k: 'scanned', label: 'Đã quét giao', group: 'gh', kind: 'int', tip: 'Số đơn shipper đã quét giao' },
      { k: 'ok', label: 'GTC', group: 'gh', kind: 'int', tip: 'Số đơn giao thành công' },
      { k: 'rateTrip', label: '% GTC chuyến', group: 'gh', kind: 'pct', good: 'high', tip: 'GTC / Đã quét giao' }
    ];
    if (c.hasTT) {
      cols.push({ k: 'ttTotal', label: 'Tổng đơn TikTok', group: 'tt', kind: 'int', tip: 'Số đơn TikTok shipper đã quét giao' });
      cols.push({ k: 'ttOk', label: 'GTC TikTok', group: 'tt', kind: 'int', tip: 'Đơn TikTok giao thành công' });
      cols.push({ k: 'rateTT', label: '% GTC TikTok', group: 'tt', kind: 'pct', good: 'high', tip: 'GTC TikTok / Tổng đơn TikTok' });
    }
    return cols;
  }

  function sortRows(rows, key, dir) {
    if (!key) return rows;
    const isN = v => typeof v === 'number' && Number.isFinite(v);
    return [...rows].sort((a, b) => {
      const x = a[key], y = b[key];
      if (isN(x) && isN(y)) return (x - y) * dir;
      if (isN(x)) return -1;
      if (isN(y)) return 1;
      return String(x ?? '').localeCompare(String(y ?? ''), 'vi') * dir;
    });
  }

  // ---- Lịch sử theo ngày ----
  // KPI POD2400: khu vực và từng bưu cục; đạt = xanh, thấp hơn dưới 15 điểm = vàng, từ 15 điểm = đỏ
  let KPI_REGION_NAME = '', KPI_REGION = NaN; const KPI_GAP = 15;   // bản chia sẻ: không có KPI sẵn, người dùng tự nhập ở Cài đặt (thanh KPI ẩn khi chưa nhập)
  let KPI_OFFICE = {};
  try { const k = JSON.parse(localStorage.getItem('b8s_kpi') || 'null'); if (k && typeof k === 'object') { KPI_REGION_NAME = String(k.n || ''); KPI_REGION = Number.isFinite(k.r) ? k.r : NaN; if (k.o && typeof k.o === 'object') for (const [n, v] of Object.entries(k.o)) if (Number.isFinite(v)) KPI_OFFICE[n] = v; } } catch { /* bỏ qua */ }
  const kpiOff = o => { if (Number.isFinite(KPI_OFFICE[o])) return KPI_OFFICE[o]; const p = plain(o), k = Object.keys(KPI_OFFICE).find(x => plain(x) === p); return k ? KPI_OFFICE[k] : NaN; };
  const kpiLevel = (v, k) => !Number.isFinite(v) || !Number.isFinite(k) ? '' : v >= k ? 'g' : v >= k - KPI_GAP ? 'a' : 'r';
  const kpiNeed = (pair, k) => Array.isArray(pair) && Number.isFinite(pair[0]) && pair[1] > 0 ? Math.max(0, Math.ceil(k / 100 * pair[1] - pair[0] - 1e-9)) : null;
  const OFFICES = [];   // bản chia sẻ: không có bưu cục sẵn, thêm ở Cài đặt
  const WEEKDAY = ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'];
  const sameOffice = (a, b) => { const x = plain(a), y = plain(b); return !!x && !!y && (x === y || x.includes(y) || y.includes(x)); };
  const canonOffice = (name, list) => { const n = String(name ?? '').trim(); return (list || []).find(o => sameOffice(o, n)) || n; };

  const dateKey = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const keyToUTC = k => { const [y, m, d] = String(k).split('-').map(Number); return Date.UTC(y, m - 1, d); };
  const utcToKey = t => { const d = new Date(t); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; };
  const addDays = (k, n) => utcToKey(keyToUTC(k) + n * 86400000);
  const fmtDM = k => `${k.slice(8, 10)}/${k.slice(5, 7)}`;
  const fmtDMY = k => `${fmtDM(k)}/${k.slice(0, 4)}`;
  const fmtAvg = v => Number.isInteger(v) ? fmt(v) : round1(v).toLocaleString('vi-VN');

  // Tuần ISO (thứ Hai là ngày đầu tuần)
  function weekOf(k) {
    const t = keyToUTC(k), dow = (new Date(t).getUTCDay() + 6) % 7;
    const start = utcToKey(t - dow * 86400000), end = addDays(start, 6);
    const thu = keyToUTC(start) + 3 * 86400000, y = new Date(thu).getUTCFullYear();
    const doy = Math.round((thu - Date.UTC(y, 0, 1)) / 86400000) + 1, wk = Math.ceil(doy / 7);
    return { key: `${y}-W${pad(wk)}`, start, end, label: `Tuần ${wk} (${fmtDM(start)}–${fmtDM(end)}/${end.slice(0, 4)})` };
  }
  function monthBounds(k) {
    const y = +k.slice(0, 4), m = +k.slice(5, 7);
    return { first: `${k.slice(0, 7)}-01`, last: addDays(utcToKey(Date.UTC(y, m, 1)), -1) };
  }
  function periodOf(k, mode) {
    if (mode === 'week') return weekOf(k);
    if (mode === 'month') return { key: k.slice(0, 7), label: `Tháng ${Number(k.slice(5, 7))}/${k.slice(0, 4)}` };
    return { key: k, label: `${WEEKDAY[new Date(keyToUTC(k)).getUTCDay()]} ${fmtDMY(k)}` };
  }
  function rangeFor(preset, today) {
    switch (preset) {
      case 'd7': return { from: addDays(today, -6), to: today };
      case 'd30': return { from: addDays(today, -29), to: today };
      case 'wk': { const w = weekOf(today); return { from: w.start, to: w.end }; }
      case 'lastwk': { const w = weekOf(addDays(today, -7)); return { from: w.start, to: w.end }; }
      case 'mo': { const b = monthBounds(today); return { from: b.first, to: b.last }; }
      case 'lastmo': { const b = monthBounds(addDays(monthBounds(today).first, -1)); return { from: b.first, to: b.last }; }
      default: return { from: '0000-01-01', to: '9999-12-31' };
    }
  }

  const nz = v => (typeof v === 'number' && Number.isFinite(v)) ? v : 0;
  const nn = v => (typeof v === 'number' && Number.isFinite(v)) ? v : null;
  const cntOf = c => { c = c || {}; return { total: nz(c.total), unscanned: nz(c.unscanned), ok: nz(c.ok), newTotal: nz(c.newTotal), newOk: nz(c.newOk), newUns: nn(c.newUns) }; };
  const blOf = b => b ? { total: nz(b.total), wait: nz(b.wait), fail: nz(b.fail), unscanned: nz(b.unscanned), d5: nn(b.d5) } : null;
  // Ship không tính vào danh sách ship của bưu cục (cấu hình trong code, không hiển thị ra giao diện): [bưu cục, tên ship]
  const HIDE_SHIPS = [];
  function shipSkipFn(office) {
    const o = plain(canonOffice(office, OFFICES) || office || '');
    const names = HIDE_SHIPS.filter(h => plain(h[0]) === o).map(h => normName(h[1]));
    return names.length ? (label => names.includes(shipKeys(label).name)) : null;
  }
  const shipsOf = list => (Array.isArray(list) ? list : []).map(s => ({
    n: String(s.n ?? s.name ?? '').trim().slice(0, 80), sc: nz(s.sc ?? s.scanned), ok: nz(s.ok),
    tt: nz(s.tt ?? s.ttTotal), tk: nz(s.tk ?? s.ttOk), sg: nn(s.sg ?? s.signed), ft: nn(s.ft)
  })).filter(s => s.n);

  // Làm sạch một bản ghi (dùng khi đọc từ kho, khi nạp file sao lưu); trả về null nếu không hợp lệ
  function normalizeRecord(r) {
    if (!r || typeof r !== 'object') return null;
    const office = String(r.office ?? '').trim(), date = String(r.date ?? '');
    if (!office || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return null;
    const skipF = shipSkipFn(office), ships = shipsOf(r.ships).filter(x => !skipF || !skipF(x.n));
    // Số shipper có quét giao hàng (đã quét phát ít nhất 1 đơn); bản ghi cũ chưa có thì suy ra từ danh sách shipper
    const nScan = nn(r.nScan) ?? (ships.length ? ships.filter(x => x.sc > 0).length : null);
    return {
      id: `${office}|${date}`, office, date, savedAt: String(r.savedAt ?? ''), v: 1, verified: !!r.verified,
      all: cntOf(r.all), tt: cntOf(r.tt), backlog: blOf(r.backlog), odrTT: nn(r.odrTT),
      odr2: (r.odr2 && typeof r.odr2 === 'object') ? { due: nz(r.odr2.due), ok: nz(r.odr2.ok) } : null, firstOk: nn(r.firstOk),
      nWork: nz(r.nWork), nScan, workersEstimated: !!r.workersEstimated, ships,
      rb: r.rb ? String(r.rb).slice(0, 4) : '', rbAt: r.rbAt ? String(r.rbAt) : ''   // rb: 're' dựng lại từ danh sách đơn, 'hi' từ danh sách đơn của Kanban
    };
  }

  // Bản chụp chỉ số của một bưu cục trong một ngày (không có mã vận đơn, không có thông tin người nhận)
  function buildRecord({ office, now, calc, verified }) {
    const e = calc.extra || {};
    return normalizeRecord({
      office, date: dateKey(now), savedAt: now.toISOString(), verified,
      all: calc.cards.all, tt: calc.cards.tt, backlog: e.backlog || null, odrTT: e.odrTT, odr2: e.odr2 || null, firstOk: calc.firstOk ?? null,
      nWork: calc.nWork, workersEstimated: calc.workersEstimated, ships: calc.shippers
    });
  }

  // Chi tiết kèm thông báo lệch tổng đơn: số dòng API trả, số trang đã gọi, dòng trùng mã, dòng mã rỗng
  function diagText(d) {
    if (!d || typeof d !== 'object') return '';
    const p = [];
    if (Number.isFinite(d.rows)) p.push(`API trả ${d.rows} dòng${Number.isFinite(d.pages) ? ` qua ${d.pages} lần gọi` : ''}`);
    if (Number.isFinite(d.dup)) p.push(`${d.dup} dòng trùng mã`);
    if (Number.isFinite(d.empty)) p.push(`${d.empty} dòng mã rỗng`);
    if (Number.isFinite(d.dropped) && d.dropped > 0) p.push(`bỏ ${d.dropped} dòng giống hệt khi gộp trang`);
    return p.length ? ` (${p.join('; ')})` : '';
  }
  // Có nên tự lưu không: đúng bưu cục và tổng đơn khớp trang
  function decideSave(x) {
    if (!x.hasOrders) return { save: false, force: false, reason: 'Chưa đọc được bảng chi tiết đơn.' };
    const page = String(x.pageOffice ?? '').trim();
    if (!page) return { save: false, force: false, reason: 'Không đọc được tên bưu cục trên trang nên không biết lưu vào bưu cục nào.' };
    if (x.startOffice && !sameOffice(x.startOffice, page))
      return { save: false, force: true, reason: `Bưu cục trên trang đổi từ “${x.startOffice}” sang “${page}” trong lúc tải dữ liệu.` };
    if (x.expectedOffice && !sameOffice(x.expectedOffice, page))
      return { save: false, force: true, reason: `Trang đang ở bưu cục “${page}”, không phải “${x.expectedOffice}” bạn đã chọn.` };
    const hasOv = Number.isFinite(x.ovNeed);
    if (hasOv && x.ovNeed !== x.computedTotal)
      return { save: false, force: true, reason: `Tổng đơn tự đếm ${x.computedTotal} khác số trên trang ${x.ovNeed}${diagText(x.diag)}.` };
    return { save: true, verified: hasOv };
  }

  // Gộp nhiều bản ghi (nhiều ngày và/hoặc nhiều bưu cục): cộng số đếm, rồi tính lại tỷ lệ từ tổng
  function aggregate(recs) {
    const dates = [...new Set(recs.map(r => r.date))].sort(), days = dates.length;
    const A = { total: 0, unscanned: 0, ok: 0, newTotal: 0, newOk: 0 }, T = { ...A };
    let nuU = 0, nuT = 0, bl = 0, blN = 0, od = 0, odN = 0, wk = 0, sn = 0, snN = 0, b5 = 0, b5N = 0, o2d = 0, o2o = 0, o2N = 0, fo = null;
    for (const r of recs) {
      for (const k of Object.keys(A)) { A[k] += nz(r.all[k]); T[k] += nz(r.tt[k]); }
      if (Number.isFinite(r.all.newUns)) { nuU += r.all.newUns; nuT += nz(r.all.newTotal); }   // bản ghi cũ chưa có số này thì bỏ qua
      if (r.backlog) { bl += r.backlog.total; blN++; if (Number.isFinite(r.backlog.d5)) { b5 += r.backlog.d5; b5N++; } }
      if (r.odr2) { o2d += r.odr2.due; o2o += r.odr2.ok; o2N++; }
      if (Number.isFinite(r.firstOk)) fo = fo === null ? r.firstOk : Math.min(fo, r.firstOk);
      if (Number.isFinite(r.odrTT)) { od += r.odrTT; odN++; }
      wk += nz(r.nWork);
      if (Number.isFinite(r.nScan)) { sn += r.nScan; snN++; }
    }
    const lastDate = dates[days - 1] || '', last = recs.filter(r => r.date === lastDate);
    const lastBl = last.some(r => r.backlog) ? last.reduce((a, r) => a + (r.backlog ? r.backlog.total : 0), 0) : null;
    const lastOd = last.some(r => Number.isFinite(r.odrTT)) ? last.reduce((a, r) => a + (Number.isFinite(r.odrTT) ? r.odrTT : 0), 0) : null;
    return { days, A, T, bl: blN ? bl : null, od: odN ? od : null, wk, sn: snN ? sn : null, lastBl, lastOd, firstDate: dates[0] || '', lastDate, b5: b5N ? b5 : null, od2: o2N ? [o2o, o2d] : null, fo, nuU, nuT, rb: recs.some(r => r.rb) };
  }
  function metricsOf(g) {
    const d = g.days || 0, A = g.A, T = g.T;
    const avg = v => (v === null || v === undefined || !d) ? null : v / d;
    return {
      days: d, total: avg(A.total), unscanned: g.rb ? null : avg(A.unscanned),   // mốc dựng lại: không biết đơn nào đã quét vào lúc đó
      scanRate: g.rb ? null : ratio(A.total - A.unscanned, A.total),
      pod: ratio(A.newOk, A.newTotal), podN: [A.newOk, A.newTotal],
      newScan: g.nuT > 0 && !g.rb ? ratio(g.nuT - g.nuU, g.nuT) : null,
      podTT: ratio(T.newOk, T.newTotal), podTTN: [T.newOk, T.newTotal],
      gtcTrip: g.rb ? null : ratio(A.ok, A.total - A.unscanned), gtcTotal: ratio(A.ok, A.total), gtcN: [A.ok, A.total],
      backlog: avg(g.bl), backlogLast: g.lastBl, odr: avg(g.od), odrLast: g.lastOd, nWork: avg(g.wk), nScan: avg(g.sn),
      backlog5: avg(g.b5), odrPct: g.od2 ? ratio(g.od2[0], g.od2[1]) : null, odrPctN: g.od2, firstOk: g.fo
    };
  }
  function groupRecords(recs, mode) {
    const map = new Map();
    for (const r of [...recs].sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0)) {
      const p = periodOf(r.date, mode);
      let g = map.get(p.key);
      if (!g) { g = { key: p.key, label: p.label, recs: [] }; map.set(p.key, g); }
      g.recs.push(r);
    }
    return [...map.values()].map(g => { g.agg = aggregate(g.recs); g.m = metricsOf(g.agg); return g; });
  }
  // Mức tăng/giảm (điểm %) so với kỳ có dữ liệu liền trước; list sắp xếp từ cũ đến mới
  function withDeltas(list) {
    const keys = ['scanRate', 'pod', 'podTT', 'gtcTrip', 'gtcTotal'];
    list.forEach((g, i) => {
      g.delta = {};
      if (!i) return;
      const p = list[i - 1].m;
      for (const k of keys) if (Number.isFinite(g.m[k]) && Number.isFinite(p[k])) g.delta[k] = g.m[k] - p[k];
    });
    return list;
  }

  // Gộp shipper theo mã (hoặc tên) qua nhiều ngày; recs phải sắp xếp từ cũ đến mới
  function aggregateShippers(recs) {
    const map = new Map();
    for (const r of recs) for (const s of r.ships) {
      const k = shipKeys(s.n), key = k.code || k.name || s.n;
      let o = map.get(key);
      if (!o) { o = { name: s.n, offices: new Set(), dates: new Set(), sc: 0, ok: 0, tt: 0, tk: 0, sg: 0, sgN: 0 }; map.set(key, o); }
      o.name = s.n; o.offices.add(r.office); o.dates.add(r.date);
      o.sc += s.sc; o.ok += s.ok; o.tt += s.tt; o.tk += s.tk;
      if (s.sg !== null) { o.sg += s.sg; o.sgN++; }
    }
    return [...map.entries()].map(([key, o]) => ({
      key, name: o.name, office: [...o.offices].join(', '), days: o.dates.size,
      sc: o.sc, avgSc: o.dates.size ? o.sc / o.dates.size : null, ok: o.ok, rateTrip: ratio(o.ok, o.sc),
      tt: o.tt, tk: o.tk, rateTT: ratio(o.tk, o.tt), sg: o.sgN ? o.sg : null
    }));
  }

  // Sao lưu / khôi phục
  function validateBackup(obj) {
    if (!obj || typeof obj !== 'object' || obj.app !== '800best-history' || !Array.isArray(obj.records))
      return { ok: false, records: [], error: 'File không phải bản sao lưu lịch sử 800BEST.' };
    const records = obj.records.map(normalizeRecord).filter(Boolean);
    if (!records.length) return { ok: false, records: [], error: 'File sao lưu không có bản ghi hợp lệ.' };
    return { ok: true, records, skipped: obj.records.length - records.length };
  }
  function planMerge(existing, incoming) {
    const put = []; let added = 0, updated = 0, kept = 0;
    for (const r of incoming) {
      const cur = existing.get(r.id);
      if (!cur) { put.push(r); added++; }
      else if (String(r.savedAt) > String(cur.savedAt)) { put.push(r); updated++; }
      else kept++;
    }
    return { put, added, updated, kept };
  }
  const DAY_MS = 86400000;
  function backupDue(count, lastMs, nowMs) { return count >= 3 && (!lastMs || nowMs - lastMs > 7 * DAY_MS); }

  /* ---- ANM: so sánh nhiều bưu cục ---- */
  const lastDayOfMonth = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate(); // m = 1..12
  function sameDayPrevMonth(k) {
    let y = +k.slice(0, 4), m = +k.slice(5, 7) - 1; const d = +k.slice(8, 10);
    if (!m) { m = 12; y--; }
    return `${y}-${pad(m)}-${pad(Math.min(d, lastDayOfMonth(y, m)))}`;
  }
  // Ngày gần target nhất (trong ±3 ngày) có cùng thứ với ref
  function nearestSameWeekday(target, ref) {
    const dt = (new Date(keyToUTC(ref)).getUTCDay() - new Date(keyToUTC(target)).getUTCDay() + 7) % 7;
    return addDays(target, dt > 3 ? dt - 7 : dt);
  }
  const CMP_MODES = ['slot2', 'yd', 'wk1', 'mo_dom', 'mo_dow', 'week', 'month'];
  function cmpPeriods(base, mode) {
    const day = k => ({ from: k, to: k, label: periodOf(k, 'day').label });
    switch (mode) {
      case 'slot2': case 'yd': return { cur: day(base), prev: day(addDays(base, -1)) };
      case 'mo_dom': return { cur: day(base), prev: day(sameDayPrevMonth(base)) };
      case 'mo_dow': return { cur: day(base), prev: day(nearestSameWeekday(sameDayPrevMonth(base), base)) };
      case 'week': {
        const w = weekOf(base), ps = addDays(w.start, -7), pe = addDays(base, -7);
        return { cur: { from: w.start, to: base, label: `Tuần này ${fmtDM(w.start)}–${fmtDM(base)}` }, prev: { from: ps, to: pe, label: `Tuần trước ${fmtDM(ps)}–${fmtDM(pe)}` } };
      }
      case 'month': {
        const pe = sameDayPrevMonth(base), first = `${base.slice(0, 7)}-01`, pf = `${pe.slice(0, 7)}-01`;
        return {
          cur: { from: first, to: base, label: `Tháng ${+base.slice(5, 7)}/${base.slice(0, 4)} (1–${+base.slice(8, 10)})` },
          prev: { from: pf, to: pe, label: `Tháng ${+pe.slice(5, 7)}/${pe.slice(0, 4)} (1–${+pe.slice(8, 10)})` }
        };
      }
      default: return { cur: day(base), prev: day(addDays(base, -7)) };
    }
  }

  const ANM_METRICS = [
    { k: 'total', label: 'Tổng đơn', group: 'Sản lượng', kind: 'avg' },
    { k: 'perNV', label: 'Đơn/NV', group: 'Sản lượng', kind: 'avg' },
    { k: 'nWork', label: 'DM đi làm', group: 'Sản lượng', kind: 'avg' },
    { k: 'nScan', label: 'DM có quét giao', group: 'Sản lượng', kind: 'avg', high: true },
    { k: 'scanRate', label: 'Tỷ lệ quét giao', group: 'Quét giao', kind: 'pct' },
    { k: 'unscanned', label: 'Đơn chưa quét', group: 'Quét giao', kind: 'avg', low: true },
    { k: 'pod', label: 'POD2400', group: 'Giao thành công', kind: 'pct' },
    { k: 'gtcTotal', label: 'GTC Total', group: 'Giao thành công', kind: 'pct' },
    { k: 'backlog', label: 'Backlog 3 ngày', group: 'Tồn đọng', kind: 'avg', low: true },
    { k: 'podTT', label: 'POD TikTok', group: 'TikTok Shop', kind: 'pct' },
    { k: 'gtcTT', label: 'GTC TikTok', group: 'TikTok Shop', kind: 'pct' },
    { k: 'odr', label: 'TikTok đến hạn ODR', group: 'TikTok Shop', kind: 'avg', low: true },
    { k: 'odrPct', label: 'Tỷ lệ ODR', group: 'TikTok Shop', kind: 'pct' },
    { k: 'backlog5', label: 'Backlog 5 ngày', group: 'Tồn đọng', kind: 'avg', low: true }
  ];
  // Bốn tiêu chí xếp hạng thi đua giữa các bưu cục
  const RANK_KEYS = ['pod', 'podTT', 'gtcTotal', 'odrPct'];
  const RANK_LABEL = { pod: 'POD2400', podTT: 'POD2400 TikTok', gtcTotal: 'Giao thành công', odrPct: 'Tỷ lệ ODR' };
  // Ngưỡng mặc định: chỉ số tỷ lệ (cao là tốt): dưới lo = đỏ, từ hi = xanh; chỉ số đếm (thấp là tốt): ≤ ok xanh, > warn đỏ
  const ANM_TH = {
    scanRate: { lo: 85, hi: 92 }, pod: { lo: 60, hi: 75 }, gtcTotal: { lo: 70, hi: 80 },
    podTT: { lo: 60, hi: 70 }, gtcTT: { lo: 60, hi: 70 },
    backlog: { ok: 50, warn: 100 }, odr: { ok: 20, warn: 40 }, odrPct: { lo: 70, hi: 85 }, backlog5: { ok: 30, warn: 80 }
  };
  function levelOf(k, v, user) {
    const t = Object.assign({}, ANM_TH[k], user && user[k]);
    if (!ANM_TH[k] || !Number.isFinite(v)) return '';
    if (t.lo !== undefined) return v >= t.hi ? 'g' : v >= t.lo ? 'a' : 'r';
    return v <= t.ok ? 'g' : v <= t.warn ? 'a' : 'r';
  }
  // Chỉ số của một nhóm bản ghi (một hoặc nhiều ngày/bưu cục); null nếu không có bản ghi
  function anmMetrics(recs) {
    if (!recs || !recs.length) return null;
    const g = aggregate(recs), m = metricsOf(g);
    m.perNV = g.wk > 0 ? g.A.total / g.wk : null;
    m.gtcTT = ratio(g.T.ok, g.T.total);
    m.n = recs.length;
    // Bản ghi so sánh dựng từ Kanban: các chỉ số quét giao / GTC chỉ có số CUỐI NGÀY (lấy từ số tổng của Kanban), tính lại từ tổng
    if (recs.every(r => r && r.fin)) {
      let T = 0, U = 0, O = 0, NT = 0, NU = 0, uOk = true, nOk = true;
      for (const r of recs) { const f = r.fin; T += nz(f.total); O += nz(f.ok); NT += nz(f.newTotal); if (Number.isFinite(f.unscanned)) U += f.unscanned; else uOk = false; if (Number.isFinite(f.newUns)) NU += f.newUns; else nOk = false; }
      m.unscanned = uOk ? U / Math.max(1, m.days) : null; m.scanRate = uOk ? ratio(T - U, T) : null; m.gtcTrip = uOk ? ratio(O, T - U) : null;
      m.gtcTotal = ratio(O, T); m.gtcN = [O, T]; m.newScan = nOk ? ratio(NT - NU, NT) : null; m.fin = true;
    }
    return m;
  }
  // Các chỉ số quét giao / GTC chỉ so với số cuối ngày chốt (không so theo giờ, không có mũi tên)
  const FIN_KEYS = ['scanRate', 'unscanned', 'newScan', 'gtcTrip', 'gtcTotal'];
  function buildAlerts(rows, user) {
    const out = [], lab = k => ANM_METRICS.find(x => x.k === k).label;
    for (const r of rows) {
      if (!r.cur) { out.push({ level: 'a', office: r.office, text: r.noData || 'Chưa có số liệu trong kỳ đang xem (cần tải bưu cục này).' }); continue; }
      const red = new Set();
      for (const k of ['pod', 'gtcTotal', 'scanRate', 'podTT', 'gtcTT']) {
        const v = r.cur[k];
        if (Number.isFinite(v) && levelOf(k, v, user) === 'r') {
          red.add(k);
          out.push({ level: 'r', office: r.office, text: `${lab(k)} chỉ ${fmtVal(v, 'pct')} (dưới ${Object.assign({}, ANM_TH[k], user && user[k]).lo}%)` });
        }
      }
      const bl = r.cur.backlog, od = r.cur.odr;
      if (Number.isFinite(bl) && levelOf('backlog', bl, user) === 'r') { red.add('backlog'); out.push({ level: 'r', office: r.office, text: `Backlog 3 ngày ${fmtAvg(bl)} đơn (trên ${Object.assign({}, ANM_TH.backlog, user && user.backlog).warn})` }); }
      if (Number.isFinite(od) && levelOf('odr', od, user) === 'r') out.push({ level: 'r', office: r.office, text: `${fmtAvg(od)} đơn TikTok đến hạn ODR (trên ${Object.assign({}, ANM_TH.odr, user && user.odr).warn})` });
      if (r.prev) {
        for (const k of ['pod', 'gtcTotal', 'scanRate']) {
          const d = r.cur[k] - r.prev[k];
          if (!red.has(k) && Number.isFinite(d) && d <= -3) out.push({ level: 'a', office: r.office, text: `${lab(k)} giảm ${round1(-d).toLocaleString('vi-VN')} điểm so với kỳ trước` });
        }
      }
      if (r.recs && r.recs.some(x => !x.verified)) out.push({ level: 'a', office: r.office, text: 'Có số liệu chưa đối chiếu được với trang gốc (dấu ⚠ trong Lịch sử).' });
    }
    const merged = [];
    for (const lv of ['r', 'a']) for (const r of rows) {
      const t = out.filter(x => x.level === lv && x.office === r.office).map(x => x.text);
      if (t.length) merged.push({ level: lv, office: r.office, text: t.join(' · ') });
    }
    return merged;
  }

  /* ---- Bản chụp theo mốc giờ (tab ANM) ---- */
  const SNAP_SLOTS = ['23:30'];   // v3.19: chỉ còn mốc chốt cuối ngày (lưu số DM, giờ GTC đầu tiên của ngày)
  const SNAP_WINDOW_MIN = 15, SNAP_KEEP_DAYS = 90, SNAP_RETRY_MS = 170000, SNAP_MAX_TRY = 5;
  const slotMin = sl => { const m = /^(\d{2}):(\d{2})$/.exec(String(sl)); return m && +m[1] < 24 && +m[2] < 60 ? +m[1] * 60 + +m[2] : NaN; };
  const minutesOf = d => d.getHours() * 60 + d.getMinutes() + d.getSeconds() / 60;
  // Mốc đang trong khung chạy (từ đúng giờ mốc đến +15 phút); trả về mốc muộn nhất nếu có
  function dueSlot(now) {
    const t = minutesOf(now); let hit = null;
    for (const sl of SNAP_SLOTS) { const m = slotMin(sl); if (t >= m && t < m + SNAP_WINDOW_MIN) hit = sl; }
    return hit;
  }
  // Mốc gần nhất đã đến giờ trong ngày (dùng để so số mới nhất với bản lưu cùng mốc)
  function lastSlotBefore(now) { const t = minutesOf(now); let r = null; for (const sl of SNAP_SLOTS) if (t >= slotMin(sl)) r = sl; return r; }
  function slotPhase(dateStr, slot, now) {
    const today = dateKey(now);
    if (dateStr < today) return 'past';
    if (dateStr > today) return 'future';
    const t = minutesOf(now), m = slotMin(slot);
    return t < m ? 'future' : t < m + SNAP_WINDOW_MIN ? 'window' : 'past';
  }
  // Bản chụp: cùng cấu trúc bản ghi ngày + mốc giờ (không có mã vận đơn, không có thông tin người nhận)
  function normalizeSnap(r) {
    const base = normalizeRecord(r), slot = String((r && r.slot) ?? '');
    if (!base || !Number.isFinite(slotMin(slot))) return null;
    // Bản chụp giữ danh sách ship dạng gọn (tên, số quét, số giao thành công, giờ quét đầu) để so sánh từng ship cùng mốc giờ
    const ships = base.ships.slice(0, 150).map(x => ({ n: x.n, sc: x.sc, ok: x.ok, tt: 0, tk: 0, sg: null, ft: x.ft }));
    return { ...base, id: `${base.office}|${base.date}|${slot}`, slot, runAt: String((r && r.runAt) ?? base.savedAt), ships };
  }
  function buildSnap({ office, now, calc, verified, slot, rb, rbAt }) {
    return normalizeSnap({ ...buildRecord({ office, now, calc, verified }), slot, runAt: now.toISOString(), rb, rbAt });
  }
  function snapsToIndex(list) {   // Map(bưu cục → Map('ngày|mốc' → bản chụp))
    const idx = new Map();
    for (const r of list) { if (!idx.has(r.office)) idx.set(r.office, new Map()); idx.get(r.office).set(`${r.date}|${r.slot}`, r); }
    return idx;
  }
  const snapAt = (idx, office, date, slot) => { const m = idx.get(office); return (m && m.get(`${date}|${slot}`)) || null; };
  function snapsToDelete(list, today, keepDays = SNAP_KEEP_DAYS) { const lim = addDays(today, -keepDays); return list.filter(r => r.date < lim).map(r => r.id); }
  // Tình trạng một mốc của một ngày: ok (đủ bưu cục) / part (thiếu) / due (đang trong khung chạy) / miss (lỡ) / future
  function slotState(dateStr, slot, now, offices, idx) {
    const have = offices.filter(o => snapAt(idx, o, dateStr, slot)).length, n = offices.length;
    if (n && have === n) return { st: 'ok', have, n };
    if (have) return { st: 'part', have, n };
    const ph = slotPhase(dateStr, slot, now);
    return { st: ph === 'future' ? 'future' : ph === 'window' ? 'due' : 'miss', have, n };
  }

  /* ---- Thời tiết (Open-Meteo, lượng mưa theo giờ) ---- */
  const WX_DEFAULT = { on: false, lat: 21.0285, lon: 105.8542 };   // bản chia sẻ: tắt sẵn, nhập tọa độ của bạn ở Cài đặt rồi bật
  const wxUrl = (lat, lon) => `https://api.open-meteo.com/v1/forecast?latitude=${encodeURIComponent(lat)}&longitude=${encodeURIComponent(lon)}&hourly=precipitation&past_days=14&forecast_days=1&timezone=Asia%2FBangkok`;
  function parseWeather(j) {
    const h = j && j.hourly;
    if (!h || !Array.isArray(h.time) || !Array.isArray(h.precipitation) || h.time.length !== h.precipitation.length || !h.time.length) return null;
    return { times: h.time.map(String), mm: h.precipitation.map(v => (typeof v === 'number' && Number.isFinite(v)) ? v : null) };
  }
  // Tổng lượng mưa (mm) của một ngày từ 06:00 đến trước phút uptoMin; null nếu không có số liệu giờ nào trong khoảng đó
  function rainSum(w, dayKey, uptoMin) {
    if (!w || !Array.isArray(w.times)) return null;
    let sum = 0, n = 0;
    for (let i = 0; i < w.times.length; i++) {
      const t = w.times[i];
      if (t.slice(0, 10) !== dayKey) continue;
      const m = +t.slice(11, 13) * 60;
      if (m >= 360 && m < uptoMin && w.mm[i] !== null) { sum += w.mm[i]; n++; }
    }
    return n ? Math.round(sum * 10) / 10 : null;
  }

  /* ---- Phân tích nguyên nhân POD2400 tăng/giảm so với hôm qua cùng mốc giờ ---- */
  const RS_TH = { minDelta: 0.5, shipPct: 10, vol: 15, load: 10, late: 30, first: 30, rain: 2, rainGap: 1.5, shipRate: 20, shipMin: 15, top: 4, arrMin: 25, freshWin: 120, freshPct: 8, oldAbs: 15, oldPct: 15, ttShare: 8, ttPod: 10, ttMin: 15, shipDrop: 5, hold: 25, curveGap: 5, curveRel: 0.15, unscan: 10 };
  const shipLabel = n => (String(n ?? '').replace(/\s*\([^)]*\)\s*$/, '').trim() || String(n ?? '').trim());
  function shipMapOf(ships) {
    const m = new Map();
    for (const x of (ships || [])) { const k = shipKeys(x.n), key = k.code || k.name || x.n; if (!m.has(key)) m.set(key, x); if (k.name && !m.has(k.name)) m.set(k.name, x); }   // cả mã lẫn tên: số liệu Kanban của hôm qua chỉ có tên ship
    return m;
  }
  const pct0 = v => `${Math.round(v)}%`;
  // v3.20: thống kê từ danh sách đơn Kanban đã nén (kbPack) tại phút m, dùng chung một cách tính cho hôm nay và hôm qua
  function kbStatsAt(E, m) {
    const P = E.P, ships = new Map(), sg = [], arrNew = [];
    let n = 0, newN = 0, newOk = 0, fresh = 0, old = 0, ttN = 0, ttOk = 0, smN = 0, smOk = 0;
    for (let i = 0; i < P.length; i += 4) {
      const arr = P[i], s = P[i + 1], f = P[i + 2], si = P[i + 3];
      if (arr !== KB_NONE && arr > m) continue;                    // chưa nhập kho tại thời điểm đó
      n++;
      const signed = s !== KB_NONE && s <= m, isNew = !!(f & KB_F.NEW), isTT = !!(f & KB_F.TT);
      if (signed) sg.push(s);
      if (isNew) {
        newN++; if (signed) newOk++;
        if (arr !== KB_NONE) { arrNew.push(arr); if (arr > m - RS_TH.freshWin) fresh++; }
        if (isTT) { ttN++; if (signed) ttOk++; } else { smN++; if (signed) smOk++; }
      } else old++;
      if (si >= 0) {
        const nm = E.ships[si]; let h = ships.get(nm);
        if (!h) ships.set(nm, h = { asg: 0, ok: 0, ft: null });
        h.asg++; if (signed) { h.ok++; if (h.ft === null || s < h.ft) h.ft = s; }
      }
    }
    arrNew.sort((x, y) => x - y);
    return { n, newN, newOk, fresh, old, ttN, ttOk, smN, smOk, ships, sg, arrMed: arrNew.length ? arrNew[arrNew.length >> 1] : null };
  }
  // Từ giờ nào đường cong ký nhận lệch so với hôm qua (so mỗi nửa giờ, và còn lệch ở mốc hiện tại); null nếu không lệch rõ
  function kbCurve(A, B, m, sign, strict) {
    const cnt = (arr, t) => { let c = 0; for (const x of arr) if (x <= t) c++; return c; };
    const gapMin = strict ? RS_TH.curveGap : 3, rel = strict ? RS_TH.curveRel : 0.08;
    const gap = t => { const ca = cnt(A.sg, t), cb = cnt(B.sg, t); return { g: (ca - cb) * sign, thr: Math.max(gapMin, rel * Math.max(ca, cb)), ca, cb }; };
    const end = gap(m); if (end.g < end.thr) return null;
    for (let t = 540; t <= m; t += 30) { const x = gap(t); if (x.g >= x.thr) return { t0: t, ca: end.ca, cb: end.cb }; }
    return { t0: m, ca: end.ca, cb: end.cb };
  }
  // cur, yd: bản ghi (ngày hoặc bản chụp) của cùng một bưu cục (kèm kbE/kbM là danh sách đơn Kanban nếu có); rain: { cur, yd } (mm) hoặc null
  function buildReasons({ cur, yd, wk, rain }) {
    if (!cur || !yd) return { dir: 'none', delta: null, reasons: [], why: !cur ? 'Chưa có số liệu của kỳ đang xem.' : 'Chưa lấy được số liệu hôm qua để so sánh (xem Cài đặt → Dữ liệu so sánh).' };
    const pc = ratio(cur.all.newOk, cur.all.newTotal), py = ratio(yd.all.newOk, yd.all.newTotal);
    if (!Number.isFinite(pc) || !Number.isFinite(py)) return { dir: 'none', delta: null, reasons: [], why: 'Chưa tính được POD2400 của một trong hai ngày.' };
    const delta = pc - py;
    if (Math.abs(delta) < RS_TH.minDelta) return { dir: 'flat', delta, pc, py, reasons: [] };
    const sign = delta > 0 ? 1 : -1, out = [], notes = [];
    const add = (s, icon, text, w, k = 1) => { if (s === sign) out.push({ icon, text, w: Math.min(5, w) * k }); };   // k: độ tin cậy của loại lý do
    // Danh sách đơn Kanban của hôm nay và hôm qua: dùng khi số đơn khớp với số của tool (Kanban cập nhật theo đợt nên có thể chậm hơn)
    const A = cur.kbE ? kbStatsAt(cur.kbE, cur.kbM) : null, B = yd.kbE ? kbStatsAt(yd.kbE, yd.kbM) : null;
    let packOk = false, signOk = false;
    if (A && B) {
      const tol = Math.max(3, 0.03 * cur.all.newTotal);
      packOk = Math.abs(A.newN - cur.all.newTotal) <= tol;
      signOk = packOk && Math.abs(A.newOk - cur.all.newOk) <= tol;
      if (!packOk) notes.push('Danh sách đơn Kanban hôm nay chưa khớp số đơn mới của tool (Kanban cập nhật chậm hơn) nên chưa phân tích theo giờ hàng về và từng ship.');
      else if (!signOk) notes.push('Số đơn ký nhận trên Kanban hôm nay chưa khớp số của tool (Kanban cập nhật chậm hơn): phần ký nhận theo ship và theo giờ dùng số liệu của tool, bỏ qua mục ship ôm đơn.');
    }
    if (!cur.kbE && yd.kbE) notes.push('Chưa có danh sách đơn hôm nay từ Kanban (đang lấy hoặc lấy lỗi) nên phân tích theo giờ hàng về và từng ship còn hạn chế.');
    // A. Hàng đầu vào
    if (packOk && A.arrMed !== null && B.arrMed !== null) {
      const d = A.arrMed - B.arrMed;
      if (Math.abs(d) >= RS_TH.arrMin) add(d > 0 ? -1 : 1, '📥', `Hàng về kho ${d > 0 ? 'trễ' : 'sớm'} hơn hôm qua: một nửa số đơn mới đã về lúc ${hmOf(A.arrMed)} (hôm qua ${hmOf(B.arrMed)})`, Math.abs(d) / RS_TH.arrMin, 1.5);
    }
    if (packOk && A.newN > 0 && B.newN > 0) {
      const fa = A.fresh / A.newN * 100, fb = B.fresh / B.newN * 100, dd = fa - fb;
      if (Math.abs(dd) >= RS_TH.freshPct) add(dd > 0 ? -1 : 1, '🕐', `${fmt(A.fresh)} đơn (${pct0(fa)}) mới về trong ${RS_TH.freshWin / 60} giờ qua, chưa kịp giao (hôm qua ${pct0(fb)})`, Math.abs(dd) / RS_TH.freshPct, 1.3);
    }
    if (yd.all.newTotal > 0) {
      const vp = (cur.all.newTotal - yd.all.newTotal) / yd.all.newTotal * 100;
      if (Math.abs(vp) >= RS_TH.vol) add(vp > 0 ? -1 : 1, '📦', `Sản lượng đơn mới ${vp > 0 ? 'tăng' : 'giảm'} ${pct0(Math.abs(vp))} (${fmt(cur.all.newTotal)} so với ${fmt(yd.all.newTotal)} đơn)`, Math.abs(vp) / RS_TH.vol, 1.4);
    }
    if (packOk) {
      const dOld = A.old - B.old, rel = B.old > 0 ? dOld / B.old * 100 : (dOld > 0 ? 100 : 0);
      if (Math.abs(dOld) >= RS_TH.oldAbs && Math.abs(rel) >= RS_TH.oldPct) add(dOld > 0 ? -1 : 1, '🗂', `Đơn tồn cũ phải phát ${dOld > 0 ? 'nhiều' : 'ít'} hơn hôm qua: ${fmt(A.old)} so với ${fmt(B.old)} đơn`, Math.abs(rel) / RS_TH.oldPct, 1.0);
    }
    if (signOk && A.newN > 0 && B.newN > 0) {
      const sA = A.ttN / A.newN * 100, sB = B.ttN / B.newN * 100, dS = sA - sB;
      const pTA = A.ttN > 0 ? A.ttOk / A.ttN * 100 : null, pSA = A.smN > 0 ? A.smOk / A.smN * 100 : null;
      if (Math.abs(dS) >= RS_TH.ttShare && pTA !== null && pSA !== null && Math.abs(pTA - pSA) >= RS_TH.ttPod) {
        add((dS > 0 ? 1 : -1) * (pTA < pSA ? -1 : 1), '🎵', `Tỉ trọng đơn TikTok ${dS > 0 ? 'tăng' : 'giảm'} từ ${pct0(sB)} ${dS > 0 ? 'lên' : 'còn'} ${pct0(sA)} (POD TikTok ${pct0(pTA)}, SME ${pct0(pSA)})`, Math.abs(dS) / RS_TH.ttShare, 0.9);
      }
      const pTB = B.ttN > 0 ? B.ttOk / B.ttN * 100 : null, pSB = B.smN > 0 ? B.smOk / B.smN * 100 : null;
      if (A.ttN >= RS_TH.ttMin && B.ttN >= RS_TH.ttMin && pTA !== null && pTB !== null && pSA !== null && pSB !== null) {
        const dT = pTA - pTB, dM = pSA - pSB;
        if (Math.abs(dT) >= RS_TH.ttPod && Math.abs(dT - dM) >= RS_TH.ttPod * 0.8) add(dT > 0 ? 1 : -1, '🎵', `POD TikTok ${dT > 0 ? 'tốt' : 'kém'} hơn hôm qua: ${pct0(pTA)} so với ${pct0(pTB)} (SME ${pct0(pSA)} so với ${pct0(pSB)})`, Math.abs(dT) / RS_TH.ttPod, 1.0);
      }
    }
    // B. Nhân lực
    const useScan = Number.isFinite(cur.nScan) && Number.isFinite(yd.nScan);
    let nc, ny, ncLabel, loadC, loadY;
    if (signOk) {
      const cs = x => { let k = 0; for (const h of x.ships.values()) if (h.ok > 0) k++; return k; };
      nc = cs(A); ny = cs(B); ncLabel = 'có đơn ký thành công'; loadC = nc > 0 ? A.newN / nc : null; loadY = ny > 0 ? B.newN / ny : null;
    } else {
      nc = useScan ? cur.nScan : cur.nWork; ny = useScan ? yd.nScan : yd.nWork; ncLabel = useScan ? 'có quét giao' : 'đi làm';
      loadC = nc > 0 ? cur.all.newTotal / nc : null; loadY = ny > 0 ? yd.all.newTotal / ny : null;
    }
    const loadPct = (loadC !== null && loadY !== null && loadY > 0) ? (loadC - loadY) / loadY * 100 : null;
    if (Number.isFinite(nc) && Number.isFinite(ny)) {
      const dN = nc - ny, rel = ny > 0 ? Math.abs(dN) / ny * 100 : 0;
      if (Math.abs(dN) >= 1 && rel >= RS_TH.shipPct) {
        const ext = (loadPct !== null && Math.abs(loadPct) >= RS_TH.load) ? `, mỗi ship ${loadPct > 0 ? 'gánh nhiều hơn' : 'nhẹ hơn'} ${pct0(Math.abs(loadPct))}` : '';
        add(dN > 0 ? 1 : -1, '🚚', `${fmt(nc)} ship ${ncLabel} (${dN > 0 ? '+' : '−'}${Math.abs(dN)} so với hôm qua)${ext}`, rel / RS_TH.shipPct, 1.6);
      } else if (loadPct !== null && Math.abs(loadPct) >= RS_TH.load) {
        add(loadPct > 0 ? -1 : 1, '🚚', `Đơn trên mỗi ship ${fmt(Math.round(loadC))} (hôm qua ${fmt(Math.round(loadY))}), ${loadPct > 0 ? 'nhiều hơn' : 'ít hơn'} ${pct0(Math.abs(loadPct))}`, Math.abs(loadPct) / RS_TH.load, 1.2);
      }
    }
    // Giờ ký đầu tiên của cả bưu cục
    if (Number.isFinite(cur.firstOk) && Number.isFinite(yd.firstOk)) {
      const d = cur.firstOk - yd.firstOk;
      if (Math.abs(d) >= RS_TH.first) add(d > 0 ? -1 : 1, '⏱', `Đơn giao thành công đầu tiên lúc ${hmOf(cur.firstOk)} (hôm qua ${hmOf(yd.firstOk)})`, Math.abs(d) / RS_TH.first, 0.6);
    }
    // Từng ship: số đơn đã ký thành công cùng giờ so với chính họ hôm qua. Hôm nay lấy từ Kanban nếu khớp, không thì từ số liệu của tool
    const keyOf = n => shipKeys(n).name || String(n ?? '').trim().toLowerCase();
    const mapT = new Map(), mapY = new Map();
    if (signOk) for (const [nm, h] of A.ships) mapT.set(keyOf(nm), { nm, ok: h.ok, asg: h.asg, ft: h.ft });
    else for (const x of (cur.ships || [])) mapT.set(keyOf(x.n), { nm: x.n, ok: nz(x.ok), asg: null, ft: null });
    if (B) for (const [nm, h] of B.ships) mapY.set(keyOf(nm), { nm, ok: h.ok, asg: h.asg, ft: h.ft });
    else for (const x of (yd.ships || [])) mapY.set(keyOf(x.n), { nm: x.n, ok: nz(x.ok), asg: null, ft: null });
    const shipData = mapT.size > 0 && mapY.size > 0;
    if (shipData) {
      const zero = [], per = [], lates = [];
      for (const k of new Set([...mapT.keys(), ...mapY.keys()])) {
        const a = mapT.get(k), b = mapY.get(k), ao = a ? a.ok : 0, bo = b ? b.ok : 0, nm = (a || b).nm;
        if (b && bo >= RS_TH.shipDrop && ao === 0) zero.push({ nm, bo });
        else if (Math.abs(ao - bo) >= RS_TH.shipDrop) per.push({ nm, ao, bo, d: ao - bo });
        if (signOk && a && b && Number.isFinite(a.ft) && Number.isFinite(b.ft)) lates.push(a.ft - b.ft);
      }
      if (sign < 0 && zero.length) {
        zero.sort((x, y) => y.bo - x.bo);
        add(-1, '🚫', `${zero.length} ship hôm qua đã ký từ ${RS_TH.shipDrop} đơn mà hôm nay chưa ký đơn nào: ${zero.slice(0, 3).map(z => `${shipLabel(z.nm)} (${z.bo})`).join(', ')}${zero.length > 3 ? '…' : ''}`, zero.reduce((q, z) => q + z.bo, 0) / 10, 1.5);
      }
      const mine = per.filter(x => (x.d > 0 ? 1 : -1) === sign).sort((x, y) => Math.abs(y.d) - Math.abs(x.d));
      mine.slice(0, 2).forEach(x => add(sign, '👤', `Ship ${shipLabel(x.nm)} ký ${fmt(x.ao)} đơn (hôm qua ${fmt(x.bo)}, cùng giờ)`, Math.abs(x.d) / RS_TH.shipDrop, 1.1));
      if (mine.length >= 3) add(sign, '👥', `${mine.length} ship ký ${sign < 0 ? 'ít' : 'nhiều'} hơn hôm qua từ ${RS_TH.shipDrop} đơn trở lên, tổng chênh ${fmt(Math.abs(mine.reduce((q, x) => q + x.d, 0)))} đơn`, mine.length / 3, 1.0);
      if (lates.length >= 3) {
        const avg = lates.reduce((q, v) => q + v, 0) / lates.length;
        if (Math.abs(avg) >= RS_TH.late) add(avg > 0 ? -1 : 1, '🕘', `${lates.length} ship ký đơn đầu tiên ${avg > 0 ? 'muộn hơn' : 'sớm hơn'} trung bình ${Math.round(Math.abs(avg))} phút so với hôm qua`, Math.abs(avg) / RS_TH.late, 0.9);
      }
      if (signOk) {
        const hold = x => { let k = 0; for (const h of x.ships.values()) if (h.asg - h.ok >= RS_TH.hold) k++; return k; };
        const hA = hold(A), hB = hold(B);
        if (hA !== hB) add(hA > hB ? -1 : 1, '📮', `${hA} ship còn ôm từ ${RS_TH.hold} đơn chưa ký trở lên (hôm qua ${hB})`, Math.abs(hA - hB), 0.8);
      }
    }
    // C. Nhịp giao: đường cong ký nhận theo giờ
    const sgA = signOk ? A.sg : (Array.isArray(cur.sgm) && Number.isFinite(cur.kbM) ? cur.sgm.filter(x => x <= cur.kbM) : null), sgB = B ? B.sg : null;
    const curve = sgA && sgB ? kbCurve({ sg: sgA }, { sg: sgB }, cur.kbM, sign, true) : null;
    const curveTxt = c => `Từ khoảng ${hmOf(c.t0)} ký ${sign < 0 ? 'chậm' : 'nhanh'} hơn hôm qua: đến ${hmOf(cur.kbM)} có ${fmt(c.ca)} đơn ký thành công (hôm qua ${fmt(c.cb)})`;
    if (curve) add(sign, '📉', curveTxt(curve), 1.5, 0.6);
    // D. Chỉ có hôm nay: đơn mới chưa quét giao
    if (sign < 0 && cur.all.newTotal > 0 && Number.isFinite(cur.all.newUns)) {
      const pu = cur.all.newUns / cur.all.newTotal * 100;
      if (pu >= RS_TH.unscan) add(-1, '📭', `Còn ${fmt(cur.all.newUns)} đơn mới chưa quét giao (${pct0(pu)} đơn mới)`, pu / RS_TH.unscan, 1.1);
    }
    // E. Thời tiết và ngày bất thường
    if (rain && Number.isFinite(rain.cur) && Number.isFinite(rain.yd)) {
      const f = v => v.toLocaleString('vi-VN');
      if (rain.cur >= RS_TH.rain && rain.cur - rain.yd >= RS_TH.rainGap) add(-1, '🌧', `Hôm nay mưa ${f(rain.cur)} mm (hôm qua ${f(rain.yd)} mm)`, (rain.cur - rain.yd) / RS_TH.rain, 0.9);
      else if (rain.yd >= RS_TH.rain && rain.yd - rain.cur >= RS_TH.rainGap) add(1, rain.cur < 0.5 ? '☀' : '🌦', `Hôm qua mưa ${f(rain.yd)} mm, hôm nay ${rain.cur < 0.5 ? 'không mưa' : f(rain.cur) + ' mm'}`, (rain.yd - rain.cur) / RS_TH.rain, 0.9);
    }
    const wdOf = d => /^\d{4}-\d{2}-\d{2}$/.test(String(d || '')) ? new Date(Date.UTC(+d.slice(0, 4), +d.slice(5, 7) - 1, +d.slice(8, 10))).getUTCDay() : null;
    const pw = wk && wk.all ? ratio(wk.all.newOk, wk.all.newTotal) : null;
    if (wdOf(cur.date) === 0) add(-1, '📅', Number.isFinite(pw) && pw < py - 3 ? `Hôm nay là Chủ nhật: Chủ nhật tuần trước cùng giờ POD cũng chỉ ${pct0(pw)} (hôm qua ${pct0(py)})` : 'Hôm nay là Chủ nhật, nhịp giao thường khác ngày thường', 3, 1.0);
    if (wdOf(yd.date) === 0) add(1, '📅', 'Hôm qua là Chủ nhật nên mức so sánh thấp hơn ngày thường', 3, 1.0);
    out.sort((a, b) => b.w - a.w);
    // Không lý do nào rõ: nêu khung giờ bắt đầu lệch (ngưỡng nhẹ hơn) thay vì bỏ trống
    if (!out.length && sgA && sgB) { const c2 = kbCurve({ sg: sgA }, { sg: sgB }, cur.kbM, sign, false); if (c2) out.push({ icon: '📉', text: curveTxt(c2), w: 0.5 }); }
    return { dir: sign > 0 ? 'up' : 'down', delta, pc, py, reasons: out.slice(0, RS_TH.top), every: out, all: out.length, shipData, packOk, signOk, notes, rainData: !!(rain && Number.isFinite(rain.cur) && Number.isFinite(rain.yd)) };
  }

  // Lỗi do phiên đăng nhập hết hạn / chưa đăng nhập
  function isAuthError(e) {
    return /HTTP\s*(401|403)\b|đăng nhập|dang nhap|login|log in|logged|unauthor|expired|hết hạn|het han|session|phiên|\btoken\b|Unexpected token|not valid JSON/i.test(String((e && e.message) ?? e ?? ''));
  }

  /* ---- Cần thêm để đạt mục tiêu (ngưỡng xanh) ---- */
  const NEED_SRC = { pod: 'podN', podTT: 'podTTN', gtcTotal: 'gtcN', odrPct: 'odrPctN' };
  // m: chỉ số của một bưu cục; trả về số đơn cần giao thành công thêm để chạm ngưỡng xanh; 0 nếu đã đạt; null nếu không tính được
  function needMore(k, m, user) {
    const src = NEED_SRC[k], pair = m && src ? m[src] : null;
    if (!Array.isArray(pair) || !Number.isFinite(pair[0]) || !(pair[1] > 0)) return null;
    const t = Object.assign({}, ANM_TH[k], user && user[k]);
    if (!Number.isFinite(t.hi)) return null;
    return Math.max(0, Math.ceil(t.hi / 100 * pair[1] - pair[0] - 1e-9));
  }

  /* ---- COD ship nộp tiền (tab ANM) ----
   * Mỗi dòng của API listCodPayment = một ship trong một ngày phát sinh (signDate).
   * Phải nộp = totalCharge, đã nộp = receivedCharge, chưa nộp = notSuccessReceivedCharge.
   * Chỉ dùng số tổng theo ship/ngày; không có mã vận đơn hay thông tin người nhận. */
  const COD_MIN = 1000, COD_LOOK = 7, COD_HALF = 0.5, COD_TODAY_ALERT_H = 23;
  // Ngày đang hiện: trước 12:00 là hôm qua (đang chốt), từ 12:00 là hôm nay (theo dõi luôn)
  function codView(now) {
    const h = now.getHours(), today = dateKey(now);
    return h < 12 ? { day: addDays(today, -1), isToday: false, track: true, today } : { day: today, isToday: true, track: true, today };
  }
  const codDateKey = s => { const m = String(s ?? '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); return m ? `${m[3]}-${pad(m[2])}-${pad(m[1])}` : ''; };
  const codName = s => String(s ?? '').replace(/\s+/g, ' ').replace(/\s*\(\s*([^)]*?)\s*\)/, ' ($1)').trim();
  function codRow(r) {
    if (!r || typeof r !== 'object') return null;
    const date = codDateKey(r.signDate), id = String(r.dispatcherId ?? r.dispatcherCode ?? '').trim();
    if (!date || !id) return null;
    const due = nn(r.totalCharge) ?? nn(r.totalCodCahrge);
    if (due === null) return null;
    const paid = nn(r.receivedCharge) ?? 0, owe = nn(r.notSuccessReceivedCharge) ?? Math.max(0, due - paid);
    const pt = nn(r.paymentTime);
    return { id, name: codName(r.dispatcherName) || id, date, due, paid, owe, cod: nn(r.totalCodCahrge), n: nn(r.signCodCount), pt: pt !== null && pt > 1e11 ? pt : null, bad: nn(r.badDebtConfirmedAmount) || 0, site: String(r.signSiteName ?? '') };
  }
  // Gộp các trang: mỗi (ship, ngày) một dòng
  function codCollect(rawRows) {
    const m = new Map();
    for (const x of rawRows || []) { const r = codRow(x); if (r) m.set(`${r.id}|${r.date}`, r); }
    return [...m.values()];
  }
  // '' = đã nộp đủ (còn nợ dưới 1.000đ); r = nộp 0đ; o = nộp dưới 50%; y = nộp từ 50% nhưng còn nợ
  function codLevel(r) {
    if (!r || !(r.owe >= COD_MIN)) return '';
    if (!(r.paid >= 1)) return 'r';
    return r.due > 0 && r.paid / r.due < COD_HALF ? 'o' : 'y';
  }
  const codAge = (today, date) => Math.round((Date.parse(`${today}T00:00:00Z`) - Date.parse(`${date}T00:00:00Z`)) / 864e5);
  // Danh sách cảnh báo THEO SHIPPER: cộng dồn COD các ngày trong 7 ngày gần nhất (hôm nay chỉ tính từ 23:00),
  // chỉ báo shipper còn chưa thu đủ. 🔴 chưa nộp đồng nào trên tổng COD đã giao; 🟠 nộp dưới 50%; 🟡 nộp từ 50% nhưng còn nợ.
  // Shipper đã nộp đủ (kể cả đã nghỉ việc) không xuất hiện. Vàng: trước 12:00 chỉ tính khi nợ có từ trước hôm qua.
  function codAlerts(byOffice, now) {
    const D = dateKey(now), Y = addDays(D, -1), h = now.getHours(), from = addDays(D, -COD_LOOK), out = [], rank = { r: 0, o: 1, y: 2 };
    for (const [office, rows] of Object.entries(byOffice || {})) {
      const g = new Map();
      for (const r of rows || []) {
        if (!r || r.date > D || r.date < from || (r.date === D && h < COD_TODAY_ALERT_H)) continue;
        const k = r.id || r.name; let x = g.get(k);
        if (!x) g.set(k, x = { office, id: r.id, name: r.name, due: 0, paid: 0, owe: 0, cod: 0, n: 0, pt: NaN, bad: 0, date: '', days: 0 });
        x.due += r.due; x.paid += r.paid; x.owe += r.owe; x.cod += r.cod || 0; x.n += r.n || 0; x.bad += r.bad || 0;
        if (Number.isFinite(r.pt) && !(r.pt <= x.pt)) x.pt = r.pt;
        if (r.owe >= COD_MIN) { x.days++; if (!x.date || r.date < x.date) x.date = r.date; }
      }
      for (const x of g.values()) {
        const lv = codLevel(x);
        if (!lv) continue;
        if (lv === 'y' && h < 12 && !(x.date && x.date < Y)) continue;
        if (lv === 'y' && x.date === D) continue;   // chỉ còn nợ của hôm nay mà đã nộp từ 50%: chưa báo
        out.push(Object.assign(x, { lv, age: codAge(D, x.date) }));
      }
    }
    return out.sort((a, b) => rank[a.lv] - rank[b.lv] || b.owe - a.owe || a.office.localeCompare(b.office));
  }
  // Tổng theo bưu cục của đúng một ngày
  function codSummary(byOffice, offices, day) {
    const by = {}, T = { due: 0, paid: 0, owe: 0, cod: 0, zero: 0, ships: 0, loaded: 0 };
    for (const o of offices || []) {
      const rows = byOffice && byOffice[o];
      if (!rows) { by[o] = null; continue; }
      const rs = rows.filter(r => r.date === day), s = { due: 0, paid: 0, owe: 0, cod: 0, zero: 0, ships: rs.length };
      for (const r of rs) { s.due += r.due; s.paid += r.paid; s.owe += r.owe; s.cod += r.cod || 0; if (codLevel(r) === 'r') s.zero++; }
      s.pct = ratio(s.paid, s.due); by[o] = s; T.loaded++;
      for (const k of ['due', 'paid', 'owe', 'cod', 'zero', 'ships']) T[k] += s[k];
    }
    T.pct = ratio(T.paid, T.due);
    return { by, tot: T };
  }
  const grp = n => String(Math.round(Math.abs(n))).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  function fmtVnd(v) {
    if (!Number.isFinite(v)) return '—';
    const a = Math.abs(v), r = Math.round(a * 100) / 100, ip = Math.floor(r), fr = Math.round((r - ip) * 100);   // đủ chữ số, không làm tròn; có số lẻ thì hiện số lẻ
    return `${v < 0 ? '-' : ''}${String(ip).replace(/\B(?=(\d{3})+(?!\d))/g, '.')}${fr ? ',' + String(fr).padStart(2, '0').replace(/0$/, '') : ''}đ`;
  }
  const fmtPct1 = v => Number.isFinite(v) ? `${v.toFixed(1).replace('.', ',')}%` : '—';
  // Giờ nộp gần nhất: “22:17 hôm qua”
  function codLastPay(pt, now) {
    if (!Number.isFinite(pt)) return 'chưa nộp lần nào';
    const t = new Date(pt), k = dateKey(t), D = dateKey(now), hm = `${pad(t.getHours())}:${pad(t.getMinutes())}`;
    return `${hm} ${k === D ? 'hôm nay' : k === addDays(D, -1) ? 'hôm qua' : fmtDM(k)}`;
  }
  const COD_ICON = { r: '🔴', o: '🟠', y: '🟡' };
  function codCopyText(alerts, now) {
    const hm = `${pad(now.getHours())}:${pad(now.getMinutes())} ${fmtDM(dateKey(now))}`;
    if (!alerts.length) return `COD ship: không có ship nào cần cảnh báo (cập nhật ${hm})`;
    return [`⚠ COD ship chưa nộp · cập nhật ${hm}`].concat(alerts.map(a =>
      `${COD_ICON[a.lv]} ${a.office} · ${a.name} · nợ ${fmtVnd(a.owe)} / ${fmtVnd(a.due)} (${a.lv === 'r' ? 'chưa nộp đồng nào' : 'đã nộp ' + fmtVnd(a.paid) + ' = ' + fmtPct1(ratio(a.paid, a.due))}) · chưa thu đủ, nợ từ ${fmtDM(a.date)}`)).join('\n');
  }

  // Ship nộp 0đ trong ngày day của các bưu cục offices (xếp theo tiền chưa nộp giảm dần)
  function codZeroList(byOffice, offices, day) {
    const out = [];
    for (const o of offices || []) for (const r of (byOffice && byOffice[o]) || []) if (r && r.date === day && codLevel(r) === 'r') out.push({ office: o, name: r.name, owe: r.owe, due: r.due });
    return out.sort((a, b) => b.owe - a.owe || a.office.localeCompare(b.office) || String(a.name).localeCompare(String(b.name)));
  }
  function codZeroText(list, label, day, now) {
    const hm = `${pad(now.getHours())}:${pad(now.getMinutes())} ${fmtDM(dateKey(now))}`;
    if (!list.length) return `COD ship nộp 0đ · ${label} · ${fmtDM(day)}: không có ship nào (cập nhật ${hm})`;
    return [`🔴 COD ship chưa nộp đồng nào · ${label} · ngày ${fmtDM(day)} · ${list.length} ship · tổng ${fmtVnd(list.reduce((t, x) => t + x.owe, 0))} · cập nhật ${hm}`].concat(list.map(x => `${x.office === label ? '' : x.office + ' · '}${x.name} · chưa nộp ${fmtVnd(x.owe)}`)).join('\n');
  }

  // Ship chưa nộp đủ trong ngày day (còn thiếu từ COD_MIN trở lên), xếp theo số tiền còn thiếu giảm dần
  function codUnpaidList(byOffice, offices, day) {
    const out = [];
    for (const o of offices || []) for (const r of (byOffice && byOffice[o]) || []) if (r && r.date === day && r.owe >= COD_MIN) out.push({ office: o, name: r.name, owe: r.owe, due: r.due, paid: r.paid });
    return out.sort((a, b) => b.owe - a.owe || a.office.localeCompare(b.office) || String(a.name).localeCompare(String(b.name)));
  }
  function codUnpaidText(list, day, now, label) {
    const hm = `${pad(now.getHours())}:${pad(now.getMinutes())} ${fmtDM(dateKey(now))}`;
    if (!list.length) return `COD ship chưa nộp đủ · ${label} · ngày ${fmtDM(day)}: không có ship nào còn thiếu (cập nhật ${hm})`;
    return [`💰 COD ship chưa nộp đủ · ${label} · ngày ${fmtDM(day)} · ${list.length} ship · còn thiếu ${fmtVnd(list.reduce((t, x) => t + x.owe, 0))} · cập nhật ${hm}`].concat(list.map(x =>
      `${x.office} · ${x.name} · còn thiếu ${fmtVnd(x.owe)} (phải nộp ${fmtVnd(x.due)}, đã nộp ${fmtVnd(x.paid)})`)).join('\n');
  }

  // ----- Tự dò bưu cục quản lý (hàm thuần) -----
  const DISC_SKIP_TYPE = ['DC', 'KA'];   // trung tâm trung chuyển (đơn gộp của các bưu cục con) và hàng cargo: không phải bưu cục giao hàng
  const discName = x => { const v = x && x.nameLocale; const t = typeof v === 'string' ? v : (v && typeof v === 'object' ? (Object.values(v).find(q => typeof q === 'string') || '') : (x && x.name) || ''); return String(t || '').trim().replace(/\s+/g, ' '); };
  // list: danh sách bưu cục hệ thống trả về; known: Set tên (đã bỏ dấu) đã có; hidden: Set tên người dùng đã xóa
  function discPick(list, known, hidden) {
    const out = { cand: [], closed: 0, skipType: [], removed: [], have: [], bad: 0, dup: 0, total: 0, allNames: new Set() };
    const seen = new Set(), seenName = new Set();
    for (const x of Array.isArray(list) ? list : []) {
      if (!x || typeof x !== 'object') { out.bad++; continue; }
      out.total++;
      const name = discName(x), id = x.id, code = x.code;
      if (!name || !id || code === undefined || code === null || code === '') { out.bad++; continue; }
      const key = String(id);
      if (seen.has(key)) { out.dup++; continue; }
      seen.add(key);
      out.allNames.add(plainOf(name));   // mọi bưu cục tài khoản này thấy được hôm nay, kể cả đang đóng cửa — dùng để phát hiện bưu cục không còn thuộc tài khoản
      if (x.status !== 'ENABLE') { out.closed++; continue; }
      if (seenName.has(name)) { out.dup++; continue; }   // tên là duy nhất; trùng tên thì chỉ lấy bưu cục đầu tiên đang mở
      seenName.add(name);
      if (DISC_SKIP_TYPE.includes(String(x.type || '').toUpperCase()) || String(x.logisticsType || '').toUpperCase() === 'CARGO') { out.skipType.push(name); continue; }
      if (known.has(plainOf(name))) { out.have.push(name); continue; }
      if (hidden.has(name) || hidden.has(plainOf(name))) { out.removed.push(name); continue; }
      out.cand.push({ name, id, code: String(code), type: String(x.type || '') });
    }
    return out;
  }
  const plainOf = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').replace(/\s+/g, ' ').trim().toUpperCase();
  // Lấy ownerSiteId từ nội dung lệnh pageQuery của trang (chuỗi JSON) – null nếu không có
  function discOwnerOf(body) {
    try { const o = typeof body === 'string' ? JSON.parse(body) : body; const v = o && o.ownerSiteId; const n = typeof v === 'string' && v.trim() ? Number(v) : v; return Number.isFinite(n) && n > 0 ? n : null; } catch { return null; }
  }
  function discSummary(pk, added, empty, failed, left, staleRemoved) {
    const parts = [];
    if (staleRemoved && staleRemoved.length) parts.push(`⚠️ loại ${staleRemoved.length} bưu cục không còn thuộc tài khoản đang đăng nhập (${staleRemoved.join(', ')})`);
    parts.push(added.length ? `thêm ${added.length} bưu cục (${added.join(', ')})` : 'không có bưu cục mới');
    if (empty.length) parts.push(`bỏ qua ${empty.length} bưu cục rỗng (${empty.join(', ')})`);
    if (pk.closed) parts.push(`${pk.closed} đã đóng`);
    if (pk.skipType.length) parts.push(`${pk.skipType.length} trung tâm/cargo (${pk.skipType.join(', ')})`);
    if (pk.removed.length) parts.push(`${pk.removed.length} đã xóa trước đó nên không thêm lại`);
    if (failed.length) parts.push(`${failed.length} chưa kiểm tra được (${failed.join(', ')})`);
    if (left) parts.push(`còn ${left} bưu cục chưa kiểm tra, bấm tải lần nữa`);
    return `Dò bưu cục: ${parts.join('; ')}.`;
  }

  // </pure>

  const $ = (s, root = document) => root.querySelector(s);
  const norm = el => (el?.innerText || el?.textContent || '').replace(/\s+/g, ' ').trim().normalize('NFC');
  const visible = el => !!(el && (el.offsetWidth || el.offsetHeight || el.getClientRects().length));
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  function log(msg, bad = false) {
    if (bad) state.lastErr = msg;
    const el = $('#b8s-log');
    if (el) {
      const line = document.createElement('div');
      if (bad) line.className = 'bad';
      line.textContent = `[${new Date().toLocaleTimeString('vi-VN')}] ${msg}`;
      el.prepend(line);
      while (el.children.length > 100) el.lastChild.remove();
    }
    const st = $('#b8s-status');
    if (st) { st.textContent = msg; st.className = bad ? 'bad' : ''; }
  }

  /* ------------------------------------------------------------------ *
   * Ghi nhận cấu trúc API (chỉ tên khóa) để phục vụ chẩn đoán
   * ------------------------------------------------------------------ */
  function listOf(payload) {
    const d = payload?.data ?? payload;
    if (Array.isArray(d)) return d;
    if (d && typeof d === 'object') for (const k of ['list', 'records', 'rows', 'items', 'resultList', 'dataList']) if (Array.isArray(d[k])) return d[k];
    return null;
  }

  function record(url, json) {
    try {
      const u = new URL(url, location.href);
      if (u.origin !== location.origin || !/\/(v9-|sea-)/.test(u.pathname)) return;
      const list = listOf(json), d = json?.data, prev = state.endpoints.get(u.pathname);
      state.endpoints.set(u.pathname, {
        count: (prev?.count || 0) + 1,
        topKeys: json && typeof json === 'object' ? Object.keys(json).slice(0, 12) : [],
        dataKeys: d && typeof d === 'object' && !Array.isArray(d) ? Object.keys(d).slice(0, 25) : undefined,
        listLength: list ? list.length : undefined,
        itemKeys: list && list[0] && typeof list[0] === 'object' ? Object.keys(list[0]).slice(0, 40) : undefined
      });
    } catch { /* bỏ qua */ }
  }

  // Bắt tiêu đề xác thực (token…) của các lệnh API trang gửi. Thư viện của trang có thể ghi đè hàm của trình duyệt,
  // nên kiểm tra định kỳ và gắn lại lên trên cùng nếu hàm của tool đã bị thay.
  function ensureHdrHook() {
    try {
      const P = XMLHttpRequest.prototype;
      if (P.setRequestHeader && P.setRequestHeader.__b8) return;
      const cur = P.setRequestHeader;
      const ours = function (k, v) {
        try { if (/^(token|lang-type|langType|TIME-ZONE|no-offset-timeZone|data-masking-N|Accept)$/i.test(k)) state.apiHdr[k] = v; } catch { /* bỏ qua */ }
        return cur.apply(this, arguments);
      };
      ours.__b8 = true;
      P.setRequestHeader = ours;
    } catch { /* bỏ qua */ }
  }

  // Xóa bưu cục/cache đã lưu trên trình duyệt này khi phát hiện đang dùng tài khoản khác chủ bưu cục khác với lần trước
  // (tránh 1 trình duyệt từng mở tài khoản A rồi mở tài khoản B vẫn còn hiện bưu cục của A)
  function resetOwnerScopedData(reason) {
    OFFICES.splice(0, OFFICES.length); offHidden.clear(); Object.keys(state.sites).forEach(k => delete state.sites[k]);
    try { localStorage.removeItem('b8s_offices'); localStorage.removeItem('b8s_offhide'); localStorage.removeItem('b8s_sites'); localStorage.removeItem('b8s_codsites'); } catch { /* bỏ qua */ }
    disc.empty = {}; disc.pick = null; disc.undo = null;
    disc.note = reason; disc.warn = true;
    try { const b = $('#b8s-of-list'); if (b) delete b.dataset.sig; } catch { /* bỏ qua */ }
    try { renderOffices(); renderOfficeList(); discRender(); if (state.view === 'anm') renderAnm(); } catch { /* bỏ qua, có thể gọi trước khi giao diện dựng xong */ }
    log(reason, true);
  }
  // Trang tự gửi lệnh danh sách bưu cục kèm ownerSiteId của chính tài khoản đang dùng: ghi lại để tự dò bưu cục
  function captureOwner(body) {
    const n = discOwnerOf(body);
    if (!n || n === state.ownerSiteId) return;
    const prev = state.ownerSiteId;
    state.ownerSiteId = n;
    try { localStorage.setItem('b8s_owner', String(n)); } catch { /* bỏ qua */ }
    if (Number.isFinite(prev) && prev > 0 && OFFICES.length) {
      resetOwnerScopedData('Phát hiện trình duyệt này vừa đổi sang tài khoản khác: đã xóa bưu cục/dữ liệu cũ. Bấm “Tải số liệu” (ANM) hoặc “Chuyển bưu cục” (Tổng quan) để tool tự tìm đúng bưu cục của tài khoản đang đăng nhập.');
    }
  }
  function hookNetwork() {
    const nativeFetch = window.fetch;
    if (typeof nativeFetch === 'function') {
      window.fetch = function (...args) {
        const p = nativeFetch.apply(this, args);
        try {
          const a = args[0];
          const url = typeof a === 'string' ? a : (a?.url || String(a || ''));
          if (/\/sea-man\/node\/pageQuery/.test(url) && args[1] && typeof args[1].body === 'string') captureOwner(args[1].body);
          p.then(res => res.clone().json().then(j => record(url, j)).catch(() => {})).catch(() => {});
        } catch { /* bỏ qua */ }
        return p;
      };
    }
    ensureHdrHook(); setInterval(ensureHdrHook, 1000);
    const xo = XMLHttpRequest.prototype.open, xs = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function (method, url, ...rest) { this.__b8surl = String(url); return xo.call(this, method, url, ...rest); };
    XMLHttpRequest.prototype.send = function (...args) {
      try { if (this.__b8surl && /\/sea-man\/node\/pageQuery/.test(this.__b8surl)) captureOwner(args[0]); } catch { /* bỏ qua */ }
      this.addEventListener('load', () => {
        try { if (this.__b8surl && (this.responseType === '' || this.responseType === 'text')) record(this.__b8surl, JSON.parse(this.responseText)); } catch { /* không phải JSON */ }
      });
      return xs.apply(this, args);
    };
  }

  /* ------------------------------------------------------------------ *
   * Đọc dữ liệu từ trang
   * ------------------------------------------------------------------ */
  // Ô “Bưu cục” trên trang là ô nhập/chọn, nên tên nằm trong value chứ không nằm trong chữ hiển thị
  function findOfficeInput() {
    const labels = [...document.querySelectorAll('label, span, div')]
      .filter(el => visible(el) && !el.closest('#b8s-app') && /^Bưu cục\s*[:：]?$/i.test(norm(el)));
    for (const el of labels) {
      let p = el;
      for (let i = 0; i < 4 && p; i++) {
        p = p.parentElement;
        const inp = p && [...p.querySelectorAll('input')].find(x => !/checkbox|radio|hidden/.test(x.type) && !x.closest('#b8s-app') && (visible(x) || x.closest('.ant-select, .el-select, .ivu-select')));
        if (inp) return inp;
      }
    }
    return null;
  }

  // Lưu ý: ô gõ của ant-select có class “ant-select-search__field” (chứa chữ “select”) nên phải loại chính nó khi tìm khung bao ngoài
  const selectWrap = inp => inp && (inp.closest('.ant-select, .el-select, .ivu-select') || (inp.parentElement && inp.parentElement.closest('[class*="select"]')) || inp.parentElement);
  function currentOfficeText(inp) {
    if (!inp) return '';
    const w = selectWrap(inp);
    // Tên đã chọn nằm ở thẻ hiển thị giá trị, không nằm trong ô gõ (ô gõ luôn trống sau khi chọn)
    const sv = w && w.querySelector && w.querySelector('.ant-select-selection-selected-value, .el-select__tags-text, .ivu-select-selected-value');
    const sel = sv && ((sv.getAttribute('title') || sv.textContent || '').trim());
    if (sel) return sel.normalize('NFC');
    let v = (inp.value || '').trim();
    if (!v) {
      v = norm(w);
      const ph = (inp.placeholder || '').trim();
      if (ph && v === ph) v = '';
      if (/^xin chọn/i.test(v)) v = '';
    }
    return v;
  }

  function readOfficeInfo() {
    const text = document.body?.innerText || '';
    const inp = findOfficeInput();
    let office = currentOfficeText(inp);
    if (!office) office = parseOfficeText(text);
    return { office, province: parseProvince(text), text };
  }

  function rememberOffice(name) {
    state.recent = [name, ...state.recent.filter(x => plain(x) !== plain(name))].slice(0, 20);
    savePrefs(); renderOfficeList();
  }

  function applyOfficeInfo({ office, province }) {
    if (province) state.province = province;
    if (!office || office === state.office) return;
    if (state.office) {
      state.emp = null; state.orders = null; state.ov = null; state.pending = null;
      log(`Đã đổi bưu cục: ${office}. Dữ liệu cũ được xóa.`);
    }
    state.office = office;
    rememberOffice(office);
  }

  // Số liệu trang tổng quan: chỉ để đối chiếu với số tự tính
  function readOverview() {
    const info = readOfficeInfo();
    applyOfficeInfo(info);
    const text = info.text;
    const get = re => { const m = text.match(re); return m ? parseNum(m[1]) : NaN; };
    const need = get(/Tổng đơn cần phát hàng\s*([\d,.]+)/i);
    const unscanned = get(/Đã nhập kho nhưng chưa quét phát\s*([\d,.]+)/i);
    if (Number.isFinite(need) || Number.isFinite(unscanned)) state.ov = { need, unscanned, at: new Date() };
  }

  const FIXED = '.el-table__fixed, .el-table__fixed-right, .vxe-table--fixed-left-wrapper, .vxe-table--fixed-right-wrapper, .fixed-left--wrapper, .fixed-right--wrapper, .ant-table-fixed-left, .ant-table-fixed-right';
  const TABLE_SEL = '.el-table, .ant-table, .vxe-table, table, [role="grid"], [role="table"]';

  function findEmployeeRoot() {
    let best = null;
    for (const r of document.querySelectorAll(TABLE_SEL)) {
      if (!visible(r) || r.closest(FIXED)) continue;
      const t = norm(r);
      if (t.includes('Nhân viên giao nhận') && t.includes('Đơn cần phát') && t.includes('Đã quét phát')) {
        if (!best || best.contains(r)) best = r;
      }
    }
    return best;
  }

  function readTable(root) {
    const headRows = [...root.querySelectorAll('thead tr')].filter(tr => !tr.closest(FIXED));
    let headCells = [];
    for (const tr of headRows) {
      const cells = [...tr.children].filter(c => /^(TH|TD)$/.test(c.tagName) && !c.classList.contains('gutter'));
      if (cells.length > headCells.length) headCells = cells;
    }
    if (!headCells.length) headCells = [...root.querySelectorAll('[role="columnheader"]')];
    const rawHeaders = headCells.map(norm);

    const rows = [];
    const trs = [...root.querySelectorAll('tbody tr')].filter(tr =>
      !tr.closest(FIXED) && !tr.classList.contains('ant-table-measure-row') && visible(tr));
    for (const tr of trs) {
      const cells = [...tr.children].filter(c => /^(TD|TH)$/.test(c.tagName)).map(norm);
      if (cells.length < 3 || !cells[0]) continue;
      if (/^(Tổng|合计|Không có dữ liệu|暂无数据)/i.test(cells[0])) continue;
      rows.push(cells);
    }
    const width = Math.max(rawHeaders.length, ...rows.map(r => r.length), 0);
    const headers = uniqueHeaders(rawHeaders, width);
    rows.forEach(r => { while (r.length < width) r.push(''); });
    return { headers, rows };
  }

  function findOrderRoot() {
    const emp = findEmployeeRoot();
    for (const r of document.querySelectorAll(TABLE_SEL)) {
      if (!visible(r) || r.closest(FIXED) || r === emp || (emp && emp.contains(r))) continue;
      const t = readTable(r);
      if (t.rows.length && waybillCol(t.headers, t.rows) >= 0) return r; // phần tử ngoài cùng (có cả tiêu đề)
    }
    return null;
  }

  /* ------------------------------------------------------------------ *
   * Gom dữ liệu qua nhiều trang, tự mở tab và quay lại tab cũ
   * ------------------------------------------------------------------ */
  const isDisabled = el => !el || el.disabled || el.hasAttribute('disabled') ||
    el.getAttribute('aria-disabled') === 'true' || /\b(disabled|ant-pagination-disabled)\b/.test(el.className || '');

  function findPager(root) {
    const rr = root.getBoundingClientRect();
    const list = [...document.querySelectorAll('.el-pagination, .ant-pagination')].filter(visible);
    const below = list.filter(p => p.getBoundingClientRect().top >= rr.top);
    const pool = below.length ? below : list;
    pool.sort((a, b) => Math.abs(a.getBoundingClientRect().top - rr.bottom) - Math.abs(b.getBoundingClientRect().top - rr.bottom));
    return pool[0] || null;
  }

  // Chữ ký nhẹ: số dòng + nội dung dòng đầu/cuối (không phân tích cả bảng)
  function tableSig(root) {
    const trs = [...root.querySelectorAll('tbody tr')].filter(tr => !tr.closest(FIXED));
    if (!trs.length) return '0';
    const t = tr => (tr.textContent || '').replace(/\s+/g, ' ').slice(0, 240);
    return trs.length + '|' + t(trs[0]) + '|' + t(trs[trs.length - 1]);
  }

  const isLoading = () => [...document.querySelectorAll('.el-loading-mask, .ant-spin-spinning')].some(visible);

  async function waitForChange(sig, finder) {
    const t0 = Date.now();
    while (Date.now() - t0 < 8000) {
      await sleep(100);
      const root = finder();
      if (root && !isLoading() && tableSig(root) !== sig) { await sleep(80); return true; }
    }
    return false;
  }

  async function goFirstPage(finder) {
    const root = finder();
    const pg = root && findPager(root);
    const first = pg && pg.querySelector('.el-pager li.number, .ant-pagination-item-1');
    if (!first || /\b(active|ant-pagination-item-active)\b/.test(first.className)) return;
    const sig = tableSig(root);
    first.click();
    await waitForChange(sig, finder);
  }

  const EMP_TAB = 'Tiến độ nhân viên';
  const isEmpTabText = t => t.normalize('NFC') === EMP_TAB;
  const isOrderTabText = t => t.length <= 40 && /^chi tiết phát/i.test(t.normalize('NFC'));

  function findTab(pred) {
    const hits = [...document.querySelectorAll('.el-tabs__item, [role="tab"], .ant-tabs-tab, li, a, span, div, button')]
      .filter(el => visible(el) && !el.closest('#b8s-app') && norm(el).length <= 60 && pred(norm(el)));
    return hits.find(el => el.matches('.el-tabs__item, [role="tab"], .ant-tabs-tab')) || hits[hits.length - 1] || null;
  }

  async function openTab(finder, pred, label) {
    if (finder()) return { switched: false };
    const target = findTab(pred);
    if (!target) { log(`Không tìm thấy tab “${label}” trên trang.`, true); return null; }
    const group = target.closest('.el-tabs__nav, .ant-tabs-nav, [role="tablist"]');
    const prev = group?.querySelector('.is-active, [aria-selected="true"], .ant-tabs-tab-active') || null;
    log(`Tự chuyển sang tab “${label}”…`);
    target.click();
    const t0 = Date.now();
    while (Date.now() - t0 < 12000) {
      await sleep(150);
      const root = finder();
      if (root && !isLoading() && root.querySelector('tbody tr')) { await sleep(150); return { switched: true, prev }; }
    }
    log(`Đã chuyển tab “${label}” nhưng chưa có dữ liệu. Nếu tab cần bấm Tìm kiếm/Truy vấn, hãy bấm trước rồi chạy lại.`, true);
    return { switched: true, prev };
  }

  async function restoreTab(h) {
    if (h && h.switched && h.prev && document.contains(h.prev)) { h.prev.click(); await sleep(200); }
  }

  // Chọn số dòng/trang lớn nhất để cần ít lần bấm trang nhất
  async function setBigPage(finder) {
    try {
      const root = finder();
      const pg = root && findPager(root);
      const box = pg && pg.querySelector('.el-pagination__sizes .el-input, .el-pagination__sizes, .ant-pagination-options-size-changer');
      if (!box) return;
      const inp = box.querySelector('input');
      const cur = inp ? parseNum(inp.value) : NaN;
      const sig = tableSig(root);
      (box.querySelector('.el-input, .ant-select-selector') || box).click();
      await sleep(250);
      const items = [...document.querySelectorAll('.el-select-dropdown__item, .ant-select-item-option')]
        .filter(el => visible(el) && !el.closest('#b8s-app'));
      let best = null, bestN = Number.isFinite(cur) ? cur : 0;
      for (const el of items) { const n = parseNum(norm(el)); if (Number.isFinite(n) && n > bestN) { best = el; bestN = n; } }
      if (!best) { inp?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true })); return; }
      best.click();
      await waitForChange(sig, finder);
      log(`Đã chọn ${bestN} dòng/trang để tải nhanh hơn.`);
    } catch { /* không đổi được thì đọc từng trang như cũ */ }
  }

  async function collectPages({ finder, pred, label, maxPages = 300, dedupe = true }) {
    let handle = null;
    try {
      handle = await openTab(finder, pred, label);
      if (!handle || !finder()) return null;
      await goFirstPage(finder);
      await setBigPage(finder);
      const seen = new Map(), raw = [];
      let headers = [], pagerTotal = 0;
      for (let page = 1; page <= maxPages; page++) {
        const root = finder();
        if (!root) { log('Mất bảng giữa chừng, dừng.', true); break; }
        const t = readTable(root);
        headers = t.headers;
        if (dedupe) t.rows.forEach(r => seen.set(JSON.stringify(r), r)); else raw.push(...t.rows);
        log(`${label}: đã đọc trang ${page} (${dedupe ? seen.size : raw.length} dòng)…`);
        const pg = findPager(root);
        const totalEl = pg && pg.querySelector('.el-pagination__total, .ant-pagination-total-text');
        if (totalEl) { const tn = parseNum(norm(totalEl)); if (Number.isFinite(tn)) pagerTotal = tn; }
        const next = pg && pg.querySelector('.btn-next, .ant-pagination-next');
        if (!next || isDisabled(next)) break;
        const sig = tableSig(root);
        next.click();
        if (!(await waitForChange(sig, finder))) { log('Trang kế tiếp không đổi nội dung, dừng.', true); break; }
      }
      await goFirstPage(finder);
      const rows = dedupe ? [...seen.values()] : raw;
      return rows.length ? { headers, rows, pagerTotal } : null;
    } finally {
      await restoreTab(handle);
    }
  }

  async function runBusy(fn) {
    if (state.busy) return false;
    state.busy = true; setBusy(true);
    try { return await fn(); }
    catch (e) { log(`Lỗi: ${e.message}`, true); return false; }
    finally { state.busy = false; setBusy(false); renderUI(); }
  }

  async function collectEmployees() {
    const res = await collectPages({ finder: findEmployeeRoot, pred: isEmpTabText, label: EMP_TAB });
    if (!res) { log('Không đọc được bảng Tiến độ nhân viên.', true); return false; }
    state.emp = { headers: res.headers, rows: res.rows, at: new Date() };
    log(`Tiến độ nhân viên: ${res.rows.length} dòng.`);
    return true;
  }

  const PII = /người nhận|nhận hàng|khách nhận|sđt|số đt|điện thoại|phone|\btel\b|địa chỉ|address|收件|电话/i;

  async function collectOrders() {
    const res = await collectPages({ finder: findOrderRoot, pred: isOrderTabText, label: 'Chi tiết phát kiện', maxPages: 500, dedupe: false });
    if (!res) { log('Không đọc được bảng Chi tiết phát kiện. Mở Cài đặt > Chẩn đoán API và gửi kết quả để chỉnh.', true); return false; }
    const wb = waybillCol(res.headers, res.rows);
    if (wb < 0) { log('Không nhận diện được cột mã vận đơn.', true); return false; }

    // Mỗi vận đơn chỉ tính một lần (giữ dòng đầy đủ nhất)
    const filled = r => r.filter(v => String(v ?? '').trim() !== '').length;
    const best = new Map();
    let noWb = 0;
    for (const r of res.rows) {
      const k = String(r[wb] ?? '').trim().toUpperCase();
      if (!k) { noWb++; continue; }
      const cur = best.get(k);
      if (!cur || filled(r) > filled(cur)) best.set(k, r);
    }
    const uniq = [...best.values()];
    log(`Đã đọc ${res.rows.length} dòng → ${uniq.length} vận đơn duy nhất${noWb ? ` (${noWb} dòng không có mã)` : ''}.`);
    if (res.pagerTotal) {
      if (res.pagerTotal === uniq.length) log(`Khớp với tổng ${fmt(res.pagerTotal)} của bộ phân trang.`);
      else log(`Chú ý: bộ phân trang báo ${fmt(res.pagerTotal)} nhưng đọc được ${fmt(uniq.length)} vận đơn.`, true);
    }

    // Mặc định bỏ cột thông tin người nhận ngay khi đọc
    const idx = res.headers.map((h, i) => (state.includePII || i === wb || !PII.test(h)) ? i : -1).filter(i => i >= 0);
    state.orders = {
      headers: idx.map(i => res.headers[i]),
      rows: uniq.map(r => idx.map(i => r[i])),
      wbCol: idx.indexOf(wb),
      at: new Date()
    };
    return true;
  }

  // Đọc tổng quan -> tiến độ nhân viên -> chi tiết phát kiện
  async function doUpdate(expect) {
    if (expect) applyOfficeInfo({ office: expect, province: '' });
    readOverview();
    const startOffice = expect || state.office || '';
    state.lastSave = null;
    let e = false, o = false, viaApi = false;
    // Tải nhanh bằng API (một lần gọi thay vì đọc từng trang 50 dòng); lỗi thì tự quay về cách đọc bảng
    const apiOffice = expect || readOfficeInfo().office || state.office || '';
    if (state.fast && apiOffice) {
      try {
        const r = await apiFetchOffice(apiOffice);
        const now = new Date();
        state.emp = { headers: r.t.emp.headers, rows: r.t.emp.rows, at: now };
        state.orders = { headers: r.t.orders.headers, rows: r.t.orders.rows, wbCol: r.t.orders.wbCol, at: now };
        e = o = viaApi = true;
        log(`⚡ Tải nhanh bằng API: ${fmt(r.t.orders.rows.length)} đơn, ${fmt(r.t.emp.rows.length)} nhân viên (${apiOffice}).`);
      } catch (err) { log(`Tải nhanh bằng API lỗi: ${err.message}. Chuyển sang đọc từng trang.`, true); }
    }
    if (!viaApi) { e = await collectEmployees(); o = await collectOrders(); }
    state.lastRefresh = new Date();
    if (o && e) log(viaApi ? 'Hoàn tất cập nhật (API).' : 'Hoàn tất cập nhật.');
    state.lastSave = o ? await trySaveHistory(expect, startOffice) : { saved: false, reason: 'Không đọc được bảng Chi tiết phát kiện.' };
    return o || e;
  }
  const updateAll = () => runBusy(doUpdate);

  /* ------------------------------------------------------------------ *
   * Chọn bưu cục ngay trên tool (điều khiển ô “Bưu cục” và nút “Tra cứu” của trang)
   * ------------------------------------------------------------------ */
  const OPT_SEL = '.el-select-dropdown__item, .ant-select-item-option, .el-cascader-node, .el-tree-node__content, li[role="option"], .el-autocomplete-suggestion li, .ivu-select-item, .vue-treeselect__option';
  const getOptions = () => [...document.querySelectorAll(OPT_SEL)]
    .filter(el => visible(el) && !el.closest('#b8s-app') && !/\bis-disabled\b|\bdisabled\b/.test(el.className || ''));

  // Nút “Tra cứu” có thể là button, a, div hoặc span: lấy phần tử có chữ đúng “Tra cứu” nằm cùng hàng với ô Bưu cục (ưu tiên bên phải)
  function searchCandidates(inp) {
    const w = (selectWrap(inp) || inp).getBoundingClientRect();
    const cy = (w.top + w.bottom) / 2;
    const out = [];
    for (const el of document.querySelectorAll('button, a, span, div, li, [role="button"], input[type="button"], input[type="submit"]')) {
      if (el.closest('#b8s-app') || !visible(el)) continue;
      const t = plain(el.tagName === 'INPUT' ? el.value : norm(el));
      if (t !== 'tra cuu') continue;
      const b = el.getBoundingClientRect();
      const dy = Math.abs((b.top + b.bottom) / 2 - cy);
      const dx = b.left >= w.right - 5 ? b.left - w.right : 3000 + Math.abs(b.left - w.right);
      out.push({ el, dy, dx, score: dy * 3 + dx * 0.1 });
    }
    return out.sort((x, y) => x.score - y.score);
  }
  function findSearchBtn(inp) {
    const c = searchCandidates(inp).find(x => x.dy <= 90);
    return c ? c.el : null;
  }

  function closeDropdown(inp) {
    try {
      inp.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true }));
      inp.blur();
    } catch { /* bỏ qua */ }
  }

  const fireMouse = (el, types) => types.forEach(t => el.dispatchEvent(new MouseEvent(t, { bubbles: true, cancelable: true, view: window })));

  // Các dòng của danh sách thả xuống: phần tử lá, hiện ra ngay bên dưới ô Bưu cục
  function dropdownItems(inp) {
    const r = inp.getBoundingClientRect(), w = selectWrap(inp)?.getBoundingClientRect() || r;
    const out = [];
    for (const el of document.querySelectorAll('body *')) {
      if (el.childElementCount || el.closest('#b8s-app') || !visible(el)) continue;
      const t = norm(el);
      if (!t || t.length > 80) continue;
      const b = el.getBoundingClientRect();
      if (b.top < Math.max(r.bottom, w.bottom) - 6 || b.top > r.bottom + 700) continue;
      if (b.left > w.right + 40 || b.right < w.left - 40) continue;
      out.push(el);
    }
    return out;
  }

  function scrollParentOf(el) {
    for (let p = el.parentElement; p && p !== document.body; p = p.parentElement) {
      const cs = getComputedStyle(p);
      if (/(auto|scroll)/.test(cs.overflowY) && p.scrollHeight > p.clientHeight + 8) return p;
    }
    return null;
  }

  async function selectOffice(name) {
    const target = plain(name);
    if (!target) { log('Hãy chọn một bưu cục.', true); return false; }
    const inp = findOfficeInput();
    if (!inp) { log('Không thấy ô “Bưu cục” trên trang. Hãy mở tab “Tiến độ giao hàng”.', true); return false; }
    if (plain(currentOfficeText(inp)) === target) { log('Đang ở đúng bưu cục này, chỉ cập nhật dữ liệu.'); return true; }

    // 1) mở danh sách như khi bạn bấm vào ô
    const opener = inp.closest('.ant-select-selector, .el-input, .el-select, .ivu-select') || selectWrap(inp) || inp;
    try { inp.focus(); } catch { /* bỏ qua */ }
    fireMouse(opener, ['mousedown', 'mouseup', 'click']);
    await sleep(500);

    // 2) gõ tên để lọc (ô tìm kiếm trong danh sách)
    if (!inp.readOnly) {
      try {
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(inp, name);
        inp.dispatchEvent(new Event('input', { bubbles: true }));
      } catch { /* bỏ qua */ }
    }
    log(`Đang tìm bưu cục “${name}” trong danh sách…`);

    // 3) tìm dòng đúng tên; nếu danh sách dài thì cuộn dần
    let pick = null, scroller = null, lastTop = -1, still = 0;
    const t0 = Date.now();
    while (Date.now() - t0 < 15000) {
      await sleep(350);
      const items = dropdownItems(inp);
      if (!items.length) { fireMouse(opener, ['mousedown', 'mouseup', 'click']); continue; }
      const r = pickOption(items.map(norm), name);
      if (r.index >= 0) { pick = items[r.index]; break; }
      if (r.ambiguous.length && Date.now() - t0 > 4000) {
        log(`Có nhiều bưu cục khớp: ${r.ambiguous.slice(0, 5).join(' | ')}. Hãy chọn tên đầy đủ hơn.`, true);
        closeDropdown(inp);
        return false;
      }
      scroller = scroller || scrollParentOf(items[0]);
      if (scroller) {
        scroller.scrollTop += Math.max(120, scroller.clientHeight * 0.8);
        still = scroller.scrollTop === lastTop ? still + 1 : 0;
        lastTop = scroller.scrollTop;
        if (still >= 3) break; // đã cuộn tới cuối
      }
    }
    if (!pick) {
      log(`Không tìm thấy bưu cục “${name}” trong danh sách của trang (có thể bạn không được cấp quyền bưu cục này).`, true);
      closeDropdown(inp);
      return false;
    }

    pick.scrollIntoView?.({ block: 'nearest' });
    fireMouse(pick, ['mouseover', 'mousedown', 'mouseup', 'click']);
    await sleep(700);
    if (!plain(currentOfficeText(inp)).includes(target)) log('Ô Bưu cục chưa hiển thị đúng tên đã chọn; vẫn tiếp tục tra cứu.', true);

    // 4) trang tự tải số liệu ngay khi chọn bưu cục (không cần bấm Tra cứu): chờ tải xong và ổn định
    await sleep(1200);
    let quiet = 0;
    const t1 = Date.now();
    while (Date.now() - t1 < 30000 && quiet < 3) {
      quiet = isLoading() ? 0 : quiet + 1;
      await sleep(300);
    }
    await sleep(600);
    log(`Đã chuyển sang bưu cục: ${currentOfficeText(inp)}`);
    return true;
  }

  // Tab tổng quan: nếu bật tải nhanh thì không cần điều khiển ô Bưu cục của trang, gọi API thẳng cho bưu cục được chọn
  async function apiOverview(name) {
    const r = await apiFetchOffice(name);
    applyOfficeInfo({ office: name, province: '' });
    const now = new Date();
    state.emp = { headers: r.t.emp.headers, rows: r.t.emp.rows, at: now };
    state.orders = { headers: r.t.orders.headers, rows: r.t.orders.rows, wbCol: r.t.orders.wbCol, at: now };
    state.lastRefresh = now;
    log(`⚡ Tải nhanh bằng API: ${fmt(r.t.orders.rows.length)} đơn, ${fmt(r.t.emp.rows.length)} nhân viên (${name}).`);
    state.lastSave = await apiSave(name, r.t, r.items, r.total, r.diag);
    return true;
  }
  async function switchOffice(name) {
    return runBusy(async () => {
      if (state.fast) {
        try { return await apiOverview(name); }
        catch (e) { log(`Tải nhanh bằng API lỗi: ${e.message}. Chuyển sang cách chọn ô Bưu cục trên trang.`, true); }
      }
      return (await selectOffice(name)) ? doUpdate(name) : false;
    });
  }

  // Danh sách bưu cục gợi ý: lấy từ hệ thống (nếu được) + các bưu cục đã dùng gần đây
  const asArray = obj => {
    if (Array.isArray(obj)) return obj;
    if (!obj || typeof obj !== 'object') return [];
    for (const k of ['list', 'records', 'rows', 'items', 'resultList', 'dataList', 'siteList', 'children', 'nodes']) if (Array.isArray(obj[k])) return obj[k];
    return obj.data && typeof obj.data === 'object' ? asArray(obj.data) : [];
  };

  async function apiJson(path, body) {
    const opt = body === undefined
      ? { method: 'GET', credentials: 'include', headers: { Accept: 'application/json' } }
      : { method: 'POST', credentials: 'include', headers: { Accept: 'application/json', 'Content-Type': 'application/json' }, body: JSON.stringify(body) };
    const r = await fetch(location.origin + path, opt);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const j = JSON.parse(await r.text());
    if (j && j.success === false) throw new Error(j.message || j.msg || 'API báo thất bại');
    return j;
  }

  function renderOfficeList() {
    const sel = $('#b8s-office-in');
    if (!sel) return;
    const cur = OFFICES.find(n => plain(n) === plain(state.office));
    const extra = state.office && !cur ? [state.office] : [];
    sel.innerHTML = `<option value="">${OFFICES.length || extra.length ? '— Chọn bưu cục —' : '— Chưa có bưu cục: bấm “Chuyển bưu cục” để tự tìm —'}</option>` +
      [...OFFICES, ...extra].map(n => `<option value="${esc(n)}">${esc(n)}</option>`).join('');
    sel.value = cur || (extra[0] || '');
  }

  /* ------------------------------------------------------------------ *
   * Lịch sử theo ngày (IndexedDB của trình duyệt, theo từng bưu cục)
   * ------------------------------------------------------------------ */
  // <idb>
  function makeStore(factory) {
    let dbp = null;
    const open = () => dbp || (dbp = new Promise((res, rej) => {
      const rq = factory.open('b8s-history', 1);
      rq.onupgradeneeded = () => { const db = rq.result; if (!db.objectStoreNames.contains('days')) db.createObjectStore('days', { keyPath: 'id' }); };
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => { dbp = null; rej(rq.error || new Error('Không mở được kho lưu trữ')); };
      rq.onblocked = () => { dbp = null; rej(new Error('Kho lưu trữ đang bị tab khác giữ')); };
    }));
    const tx = async (mode, fn) => {
      const db = await open();
      return new Promise((res, rej) => {
        const t = db.transaction('days', mode);
        let out; const r = fn(t.objectStore('days'));
        if (r) r.onsuccess = () => { out = r.result; };
        t.oncomplete = () => res(out);
        t.onerror = t.onabort = () => rej(t.error || new Error('Lỗi ghi kho lưu trữ'));
      });
    };
    return {
      all: () => tx('readonly', s => s.getAll()),
      put: rec => tx('readwrite', s => s.put(rec)),
      putMany: recs => tx('readwrite', s => { recs.forEach(r => s.put(r)); }),
      del: id => tx('readwrite', s => s.delete(id))
    };
  }
  // Kho riêng cho bản chụp theo mốc giờ (cơ sở dữ liệu khác, không ảnh hưởng kho lịch sử theo ngày)
  function makeSnapStore(factory) {
    let dbp = null;
    const open = () => dbp || (dbp = new Promise((res, rej) => {
      const rq = factory.open('b8s-snaps', 1);
      rq.onupgradeneeded = () => { const db = rq.result; if (!db.objectStoreNames.contains('snaps')) db.createObjectStore('snaps', { keyPath: 'id' }); };
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => { dbp = null; rej(rq.error || new Error('Không mở được kho bản chụp')); };
      rq.onblocked = () => { dbp = null; rej(new Error('Kho bản chụp đang bị tab khác giữ')); };
    }));
    const tx = async (mode, fn) => {
      const db = await open();
      return new Promise((res, rej) => {
        const t = db.transaction('snaps', mode);
        let out; const r = fn(t.objectStore('snaps'));
        if (r) r.onsuccess = () => { out = r.result; };
        t.oncomplete = () => res(out);
        t.onerror = t.onabort = () => rej(t.error || new Error('Lỗi ghi kho bản chụp'));
      });
    };
    return {
      all: () => tx('readonly', st => st.getAll()),
      putMany: recs => tx('readwrite', st => { recs.forEach(r => st.put(r)); }),
      delMany: ids => tx('readwrite', st => { ids.forEach(i => st.delete(i)); })
    };
  }
  // </idb>
  const store = (() => { try { return makeStore(indexedDB); } catch { return null; } })();

  async function loadHistory() {
    if (!store) { state.histErr = 'Trình duyệt không hỗ trợ IndexedDB nên không lưu được lịch sử.'; renderUI(); return; }
    try {
      const list = await store.all();
      state.hist = new Map();
      for (const r of list || []) { const n = normalizeRecord(r); if (n) state.hist.set(n.id, n); }
      state.histReady = true; state.histErr = '';
    } catch (e) { state.histErr = `Không đọc được lịch sử: ${e.message}`; }
    try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch { /* bỏ qua */ }
    renderUI();
  }

  async function putRecords(recs) {
    if (!store) throw new Error('Không có kho lưu trữ');
    await store.putMany(recs);
    recs.forEach(r => state.hist.set(r.id, r));
  }

  /* ------------------------------------------------------------------ *
   * Bản chụp theo mốc giờ cố định (tab ANM)
   * Đến mốc 10:30, 13:30, 15:30, 18:30, 23:30 (giờ máy) mà trang 800BEST đang mở: tự tải số liệu các bưu cục ANM bằng API
   * và lưu lại. Tải tay hay tự cập nhật khi vào tab ANM chỉ để xem, không tạo bản chụp.
   * ------------------------------------------------------------------ */
  const snapStore = (() => { try { return makeSnapStore(indexedDB); } catch { return null; } })();
  const TAB_ID = Math.random().toString(36).slice(2);
  const API_KB_DET = '/v9-order/dc/kanban/am/vn/site-kpi-kanban/get-sdisp-detail', API_KB_SUM = '/v9-order/dc/kanban/am/vn/site-kpi-kanban/get-sdisp', API_KB_SREC = '/v9-order/dc/kanban/am/get-srec-sdisp';
  const API_KB_SFT = '/v9-order/dc/kanban/am/vn/site-kpi-kanban/get-sdisp-shift-summary';   // v3.33: Tỷ lệ ODR thật (ký nhận đơn tới hạn ODR), cộng Trip 1 + Trip 2
  const KB_KEEP = 9, KB_RETRY_MS = 300000, KB_CMP_DAYS = [1, 7];
  // Dò toàn bộ cây JSON trả về, cộng totalCnt/finishedCnt của mọi mục có indicatorName chứa "ODR" (gồm cả Trip 1 + Trip 2 + ca khác nếu có)
  function sumOdrShift(d) {
    let due = 0, ok = 0, found = false;
    const walk = x => {
      if (!x || typeof x !== 'object') return;
      if (Array.isArray(x)) { x.forEach(walk); return; }
      if (/ODR/i.test(String(x.indicatorName || ''))) {
        const t = Number(x.totalCnt), f = Number(x.finishedCnt);
        if (Number.isFinite(t) && Number.isFinite(f)) { due += t; ok += f; found = true; }
      }
      for (const k of Object.keys(x)) walk(x[k]);
    };
    walk(d);
    return found ? { due, ok } : null;
  }
  // Tỷ lệ ODR thật (đơn giao xuôi ký nhận trong hạn / tổng đơn giao xuôi ký nhận) của 1 bưu cục 1 ngày, lấy trực tiếp từ Kanban (không phân biệt TikTok)
  async function fetchOdrReal(code, date) {
    const base = { siteCode: code, isSelfData: 0, isCargo: 0, isExpress: 0, nationCode: 'VN', statisticsDate: `${date} 00:00:00`, timeFrom: `${date}T00:00:00.000+07:00` };
    const d = await apiPost(API_KB_SFT, base);
    return sumOdrShift(d);
  }
  function makeKbStore(factory) {
    let dbp = null;
    const open = () => dbp || (dbp = new Promise((res, rej) => {
      const rq = factory.open('b8s-kb', 1);
      rq.onupgradeneeded = () => { const db = rq.result; if (!db.objectStoreNames.contains('days')) db.createObjectStore('days', { keyPath: 'id' }); };
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => { dbp = null; rej(rq.error || new Error('Không mở được kho số liệu so sánh')); };
      rq.onblocked = () => { dbp = null; rej(new Error('Kho số liệu so sánh đang bị tab khác giữ')); };
    }));
    const tx = async (mode, fn) => {
      const db = await open();
      return new Promise((res, rej) => {
        const t = db.transaction('days', mode); let out; const r = fn(t.objectStore('days'));
        if (r) r.onsuccess = () => { out = r.result; };
        t.oncomplete = () => res(out); t.onerror = t.onabort = () => rej(t.error || new Error('Lỗi ghi kho số liệu so sánh'));
      });
    };
    return { all: () => tx('readonly', st => st.getAll()), put: r => tx('readwrite', st => { st.put(r); }), delMany: ids => tx('readwrite', st => { ids.forEach(i => st.delete(i)); }), clear: () => tx('readwrite', st => { st.clear(); }) };
  }
  const kbStore = (() => { try { return makeKbStore(indexedDB); } catch { return null; } })();
  const kbKey = (office, date) => `${office}|${date}`;
  async function loadKb() {
    const K = state.kb;
    if (!kbStore) { K.err = 'Trình duyệt không hỗ trợ IndexedDB nên không giữ được số liệu so sánh (vẫn lấy được nhưng mỗi lần tải sẽ chậm hơn).'; K.ready = true; return; }
    try {
      const lim = addDays(dateKey(new Date()), -KB_KEEP), list = (await kbStore.all()) || [], dead = [];
      for (const r of list) { if (!r || !r.id || !Array.isArray(r.P) || !Array.isArray(r.ships)) { if (r && r.id) dead.push(r.id); continue; } if (r.date < lim) dead.push(r.id); else K.map.set(r.id, r); }
      if (dead.length) await kbStore.delMany(dead);
    } catch (e) { K.err = `Không đọc được kho số liệu so sánh: ${e.message}`; }
    K.ready = true;
  }
  async function kbFetchDay(code, date, light) {
    const base = { siteCode: code, isSelfData: 0, isCargo: 0, isExpress: 0, nationCode: 'VN', statisticsDate: `${date} 00:00:00` };
    const items = []; let size = 200, total = null;
    for (let page = 1; page <= 80; page++) {
      let d;
      try { d = await apiPost(API_KB_DET, Object.assign({}, base, { dataKey: 'allCnt', page, size })); }
      catch (e) { if (page === 1 && size > 50 && !isAuthError(e)) { size = 50; page--; continue; } throw e; }
      const list = d && Array.isArray(d.list) ? d.list : [];
      if (total === null) total = Number(d && d.total) || 0;
      items.push(...list);
      if (!list.length || items.length >= total) break;
    }
    if (light) return { items, total, sum: null };   // chỉ cần danh sách đơn (số hôm nay để phân tích nguyên nhân, không lưu)
    const sum = await apiPost(API_KB_SUM, Object.assign({}, base, { timeFrom: `${date}T00:00:00.000+07:00` }));
    let srec = null; try { srec = await apiPost(API_KB_SREC, Object.assign({}, base, { timeFrom: `${date}T00:00:00.000+07:00` })); } catch (e) { if (isAuthError(e)) throw e; }
    let odr = null; try { odr = sumOdrShift(await apiPost(API_KB_SFT, Object.assign({}, base, { timeFrom: `${date}T00:00:00.000+07:00` }))); } catch (e) { if (isAuthError(e)) throw e; }
    const g = k => ((sum && sum.groupIndicatorList) || []).find(x => x && x.dataKeyPrefix === k) || null;
    const one = k => { const x = ((sum && sum.singleIndicatorList) || []).find(y => y && y.dataKey === k); return x ? Number(x.value) : null; };
    const pod = g('2400Pod'), att = g('tryDispIncludStock'), sd = srec && srec.sdisp ? srec.sdisp : null;
    return { items, total, sum: {
      all: sum && Number.isFinite(Number(sum.topicIndicatorValue)) ? Number(sum.topicIndicatorValue) : null, newN: one('todayAddSdispCnt'),
      podOk: pod ? Number(pod.member) : null, podDen: pod ? Number(pod.denominator) : null, attOk: att ? Number(att.member) : null, attDen: att ? Number(att.denominator) : null,
      newAttempt: sd ? Number(sd.attemptCnt) : null, newCnt: sd ? Number(sd.sdispCnt) : null, success: sd ? Number(sd.successCnt) : null,
      odrOk: odr ? odr.ok : null, odrDen: odr ? odr.due : null
    } };
  }
  // Đối chiếu danh sách đơn với số tổng của Kanban; trả về danh sách điểm lệch (rỗng = khớp) và số cuối ngày cho các chỉ số quét giao
  function kbVerify(E, S) {
    const c = kbDayCheck(E), bad = [];
    if (Number.isFinite(S.all) && S.all !== c.all) bad.push(`tổng đơn ${c.all} ≠ ${S.all}`);
    const nn = Number.isFinite(S.newN) ? S.newN : S.podDen;
    if (Number.isFinite(nn) && nn !== c.newN) bad.push(`đơn mới ${c.newN} ≠ ${nn}`);
    if (Number.isFinite(S.podOk) && S.podOk !== c.newOk) bad.push(`POD2400 tử số ${c.newOk} ≠ ${S.podOk}`);
    const total = Number.isFinite(S.attDen) ? S.attDen : c.all;
    const fin = {
      total, unscanned: Number.isFinite(S.attOk) && Number.isFinite(S.attDen) ? S.attDen - S.attOk : null,
      ok: Number.isFinite(S.success) ? S.success : c.ok, newTotal: Number.isFinite(S.newCnt) ? S.newCnt : c.newN,
      newUns: Number.isFinite(S.newAttempt) && Number.isFinite(S.newCnt) ? S.newCnt - S.newAttempt : null,
      odrOk: Number.isFinite(S.odrOk) ? S.odrOk : null, odrDen: Number.isFinite(S.odrDen) ? S.odrDen : null
    };
    return { bad, fin, check: c };
  }
  // Lấy (nếu chưa có) số liệu của các ngày so sánh cho các bưu cục; chạy nền, ghi lỗi vào state.kb thay vì ném ra ngoài
  // v3.30.2: số so sánh đã lưu nghi sai thì TỰ tải lại, không cần bấm “Lấy lại ngay”:
  //  - bản lưu từ phiên bản cũ (v < 2): tải lại đúng một lần;
  //  - bản tải trong 6 giờ đầu sau khi ngày kết thúc (Kanban chưa chốt số cả ngày): tải lại sau mốc đó;
  //  - bản có từ 30 đơn trở lên mà không có đơn nào ký nhận (số cả ngày bằng 0): tải lại.
  //  - v3.33: bản lưu từ trước khi có Tỷ lệ ODR thật (chưa có fin.odrDen): tải lại đúng một lần để lấy số ODR thật.
  function kbStale(rec, date, nowMs) {
    if (!rec || rec.v !== 2) return true;
    if (!rec.fin || !Number.isFinite(rec.fin.odrDen)) return true;
    const end = kbDay0(date) + 86400000, settle = end + 6 * 3600000, f = Date.parse(rec.fetchedAt);
    if ((!Number.isFinite(f) || f < settle) && (nowMs ?? Date.now()) > settle) return true;
    try { const c = kbDayCheck(rec); if (c.all >= 30 && c.ok === 0) return true; } catch { return true; }
    return false;
  }
  async function kbEnsure(force) {
    const K = state.kb;
    if (K.running || !OFFICES.length || !state.fast) return;
    if (!K.ready) { K.running = true; try { await loadKb(); } finally { K.running = false; } }
    const now = new Date(), today = dateKey(now), jobs = [];
    for (const o of OFFICES) for (const k of KB_CMP_DAYS) {
      const date = addDays(today, -k), key = kbKey(o, date);
      if (K.map.has(key) && !force && !kbStale(K.map.get(key), date)) continue;   // số đã lưu nghi sai thì tự tải lại; “Lấy lại ngay” (force) tải lại tất cả
      if (!force && K.tried[key] && Date.now() - K.tried[key] < KB_RETRY_MS) continue;
      jobs.push({ o, date, key });
    }
    if (!jobs.length) return;
    K.running = true; K.msg = `Đang lấy số liệu hôm qua / tuần trước 0/${jobs.length}…`; if (state.view === 'anm' && !state.anm.loading) renderAnm();
    ensureHdrHook();
    const tk = parseList(state.tiktok); let p = 0, done = 0, auth = false;
    const worker = async () => {
      while (p < jobs.length && !auth) {
        const j = jobs[p++]; K.tried[j.key] = Date.now();
        try {
          const site = await resolveSite(j.o), day = await kbFetchDay(site.code, j.date);
          if (!day.items.length) throw new Error('API không trả đơn nào cho ngày này');
          const E = kbPack(day.items, tk, j.date), v = kbVerify(E, day.sum);
          K.diag[j.key] = { at: Date.now(), ok: !v.bad.length, text: v.bad.length ? `lệch: ${v.bad.join('; ')}` : `khớp Kanban (${v.check.all} đơn, ${v.check.newN} đơn mới, POD ${v.check.newOk}/${v.check.newN})` };
          if (v.bad.length) throw new Error(`lệch với số tổng Kanban (${v.bad.join('; ')}) nên không dùng`);
          const rec = { id: j.key, office: j.o, date: j.date, v: 2, fetchedAt: new Date().toISOString(), ships: E.ships, P: E.P, fin: v.fin };
          K.map.set(j.key, rec); delete K.fail[j.key];
          if (kbStore) { try { await kbStore.put(rec); } catch (e) { K.err = `Không lưu được số liệu so sánh: ${e.message}`; } }
        } catch (e) {
          if (isAuthError(e)) auth = true;
          K.fail[j.key] = e.message; if (!K.diag[j.key]) K.diag[j.key] = { at: Date.now(), ok: false, text: e.message };
          log(`Số liệu so sánh ${j.o} ${fmtDM(j.date)}: ${e.message}`, true);
        }
        done++; K.msg = `Đang lấy số liệu hôm qua / tuần trước ${done}/${jobs.length}…`;
        if (state.view === 'anm' && !state.anm.loading) renderAnm();
      }
    };
    try { await Promise.all([worker(), worker()]); }
    finally {
      K.running = false; const nf = jobs.filter(j => K.fail[j.key]).length;
      K.msg = auth ? 'Cần đăng nhập lại trang 800BEST để lấy số liệu hôm qua / tuần trước.' : nf ? `Chưa lấy được ${nf}/${jobs.length} ngày so sánh (xem Cài đặt → Dữ liệu so sánh).` : '';
      if (state.view === 'anm' && !state.anm.loading) renderAnm();
      if (state.view === 'set') renderKbInfo();
    }
  }
  // v3.20: danh sách đơn Kanban của HÔM NAY (chỉ giữ trong bộ nhớ, tối đa 8 phút) cho các bưu cục cần giải thích POD
  const KB_TODAY_TTL = 8 * 60000;
  async function kbEnsureToday(offs) {
    const K = state.kb, today = dateKey(new Date());
    if (K.todayRun || !state.fast || !OFFICES.length) return;
    const jobs = [...new Set(offs || [])].filter(o => {
      const e = K.today.get(o); if (e && e.date === today && Date.now() - e.at < KB_TODAY_TTL) return false;
      const t = K.todayTried[o]; return !(t && Date.now() - t < KB_RETRY_MS);
    });
    if (!jobs.length) return;
    K.todayRun = true; ensureHdrHook();
    const tk = parseList(state.tiktok); let p = 0, got = 0, auth = false;
    const worker = async () => {
      while (p < jobs.length && !auth) {
        const o = jobs[p++]; K.todayTried[o] = Date.now();
        try {
          const site = await resolveSite(o), day = await kbFetchDay(site.code, today, true);
          if (!day.items.length) throw new Error('API không trả đơn nào cho hôm nay');
          K.today.set(o, { date: today, at: Date.now(), E: kbPack(day.items, tk, today) }); delete K.todayTried[o]; got++;
        } catch (e) { if (isAuthError(e)) auth = true; log(`Danh sách đơn hôm nay (phân tích nguyên nhân) ${o}: ${e.message}`, true); }
      }
    };
    try { await Promise.all([worker(), worker()]); }
    finally { K.todayRun = false; if (got && state.view === 'anm' && !state.anm.loading) renderAnm(); }
  }
  // Bản ghi so sánh của một bưu cục ở một ngày, tại phút m của ngày; null nếu chưa có số liệu
  function kbRec(office, date, m) {
    const E = state.kb.map.get(kbKey(office, date));
    if (!E) return null;
    const r = kbAt(E, m, state.dm, office, date), dayRec = (state.hist.get(`${office}|${date}`) || null);
    r.kbE = E; r.kbM = m;
    if (dayRec && dayRec.verified !== false) {   // DM đã lưu theo đúng luật của tool (giờ quét đầu, chức vụ) thì dùng số đó
      if (Number.isFinite(dayRec.nWork) && !dayRec.workersEstimated) { r.nWork = dayRec.nWork; r.workersEstimated = false; }
      if (Number.isFinite(dayRec.nScan)) r.nScan = dayRec.nScan;
    }
    return r;
  }
  function renderKbInfo() {
    const box = $('#b8s-kb-st'), btn = $('#b8s-kb'); if (!box) return;
    const K = state.kb, today = dateKey(new Date());
    if (btn) { btn.disabled = !!K.running; btn.textContent = K.running ? 'Đang lấy…' : 'Lấy lại ngay'; }
    const rows = [];
    for (const o of OFFICES) for (const k of KB_CMP_DAYS) {
      const date = addDays(today, -k), key = kbKey(o, date), d = K.diag[key], have = K.map.has(key);
      rows.push(`<tr><td class="l">${esc(o)}</td><td>${esc(fmtDM(date))}</td><td>${have ? '✔' : '—'}</td><td class="l ${d && !d.ok ? 'bad' : ''}">${esc(d ? d.text : (have ? 'đã lưu từ lần trước' : 'chưa lấy'))}</td></tr>`);
    }
    box.innerHTML = `${K.msg ? `<div>${esc(K.msg)}</div>` : ''}${K.err ? `<div class="bad">${esc(K.err)}</div>` : ''}<div class="b8s-tw" style="max-height:220px;overflow:auto"><table><thead><tr><th class="l">Bưu cục</th><th>Ngày</th><th>Có</th><th class="l">Đối chiếu số tổng Kanban</th></tr></thead><tbody>${rows.join('')}</tbody></table></div>`;
  }

  // Tải một bưu cục bằng API và lưu: bản ghi ngày (như khi tải tay) + bản chụp của mốc
  const API_DET = '/v9-order/amTodayDispBoard/getTodayDetail', API_EMP = '/v9-order/amTodayDispBoard/getTodayEmployeeSum', API_SITE = '/sea-man/node/pageQuery';
  async function apiPost(path, body) {
    // Trang đăng nhập bằng cookie nên header token thường RỖNG; dùng bộ tiêu đề mặc định rồi ghi đè bằng giá trị đã bắt được từ trang
    const hdr = Object.assign({ Accept: 'application/json, text/plain, */*', 'lang-type': 'vi-VN', langType: 'vi-VN', 'TIME-ZONE': 'GMT+7', 'data-masking-N': 'mark', 'no-offset-timeZone': 'true', token: '', 'X-Requested-With': 'XMLHttpRequest' }, state.apiHdr);
    // Luôn gọi bằng tiếng Việt: tên chức vụ, trạng thái… đổi theo ngôn ngữ của header, nếu theo ngôn ngữ trang (en/zh) thì bộ lọc chức vụ lệch hết
    for (const k of Object.keys(hdr)) if (/^lang-?type$/i.test(k)) delete hdr[k];
    hdr['lang-type'] = 'vi-VN'; hdr.langType = 'vi-VN';
    for (const k of Object.keys(hdr)) if (typeof hdr[k] === 'boolean') hdr[k] = String(hdr[k]);
    hdr['Content-Type'] = 'application/json;charset=UTF-8';
    const r = await fetch(location.origin + path, { method: 'POST', credentials: 'include', headers: hdr, body: JSON.stringify(body) });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const j = JSON.parse(await r.text());
    if (!j || j.success === false || (j.success === undefined && j.code !== undefined && Number(j.code) !== 200)) throw new Error((j && (j.errMsg || j.message || j.msg)) || 'API báo thất bại');
    return j.data;
  }
  function loadSites() { try { const o = JSON.parse(localStorage.getItem('b8s_sites') || '{}'); if (o && typeof o === 'object') Object.assign(state.sites, o); } catch { /* bỏ qua */ } }
  function saveSites() { try { localStorage.setItem('b8s_sites', JSON.stringify(state.sites)); } catch { /* bỏ qua */ } }
  // ----- Danh sách bưu cục quản lý: thêm / xóa ở Cài đặt, lưu trên trình duyệt -----
  const offHidden = new Set();   // bưu cục đã xóa: ẩn cả khỏi lịch sử cũ
  function loadOffices() {
    try { const a = JSON.parse(localStorage.getItem('b8s_offices') || 'null'); if (Array.isArray(a) && a.every(x => typeof x === 'string' && x.trim())) OFFICES.splice(0, OFFICES.length, ...a.map(x => x.trim())); } catch { /* bỏ qua */ }
    try { const h = JSON.parse(localStorage.getItem('b8s_offhide') || '[]'); if (Array.isArray(h)) h.forEach(x => typeof x === 'string' && offHidden.add(x)); } catch { /* bỏ qua */ }
  }
  function saveOffices() {
    try { localStorage.setItem('b8s_offices', JSON.stringify(OFFICES)); localStorage.setItem('b8s_offhide', JSON.stringify([...offHidden])); } catch { /* bỏ qua */ }
  }
  const siteNm = x => { const v = x.nameLocale; if (typeof v === 'string') return v; if (v && typeof v === 'object') return Object.values(v).find(t => typeof t === 'string') || ''; return x.name || ''; };
  // Kiểm tra tên bưu cục phải ĐÚNG TUYỆT ĐỐI (kể cả chữ hoa, chữ thường, dấu). Trả về {ok, name, id, code} hoặc {ok:false, msg}
  async function officeLookup(typed) {
    const name = String(typed ?? '').trim().replace(/\s+/g, ' ');
    if (!name) return { ok: false, msg: 'Hãy gõ tên bưu cục.' };
    if (OFFICES.includes(name)) return { ok: false, msg: `“${name}” đã có trong danh sách.` };
    let seen = null;
    for (const auth of [true, false]) {
      let list;
      try { list = asArray(await apiPost(API_SITE, { pageNumber: 1, pageSize: 200, fuzzyNameOrCode: name, limitSiteByAuth: auth, statusList: ['ENABLE', 'SHUTDOWN'] })); }
      catch (e) { return { ok: false, msg: `Không kiểm tra được tên bưu cục (${e && e.message || e}). Thử lại sau.` }; }
      const hits = list.filter(x => siteNm(x).trim() === name);
      const hit = hits.find(x => x.status !== 'SHUTDOWN') || hits[0];
      if (hit) { seen = { hit, auth }; break; }
    }
    if (!seen) return { ok: false, msg: `Bưu cục “${name}” không tồn tại. Gõ đúng từng chữ, kể cả chữ hoa, chữ thường và dấu.` };
    const x = seen.hit;
    if (x.status === 'SHUTDOWN') return { ok: false, msg: `Bưu cục “${name}” đã đóng, không thể thêm.` };
    if (!x.id || !x.code) return { ok: false, msg: `Bưu cục “${name}” thiếu mã trong hệ thống, không thể thêm.` };
    if (!seen.auth) return { ok: false, msg: `Bưu cục “${name}” có trên hệ thống nhưng tài khoản của bạn không có quyền xem.` };
    return { ok: true, name, id: x.id, code: String(x.code) };
  }
  const offUi = { busy: false, msg: '', bad: false };
  function offSay(msg, bad) { offUi.msg = msg; offUi.bad = !!bad; const el = $('#b8s-of-st'); if (el) { el.textContent = msg; el.className = bad ? 'bad' : 'muted'; el.style.fontWeight = bad ? '700' : ''; } }
  function renderOffices() {
    const box = $('#b8s-of-list'); if (!box) return;
    const ids = codIds();
    const sig = JSON.stringify([OFFICES, OFFICES.map(o => ids[plain(o)] ?? null), offUi.busy]);
    if (box.dataset.sig !== sig) {
      box.dataset.sig = sig;
      box.innerHTML = OFFICES.map(n => { const v = ids[plain(n)], miss = !Number.isFinite(v);
        return `<div class="row" style="gap:8px;align-items:center;padding:3px 0;border-top:1px solid #eef2f7"><b style="min-width:170px">${esc(n)}</b><label class="muted">Mã COD <input type="number" data-cod="${esc(n)}" value="${miss ? '' : v}" placeholder="chưa có" style="width:120px;${miss ? 'border:1px solid #dc2626' : ''}"></label>${miss ? '<span class="bad">thiếu mã COD: bảng COD của bưu cục này sẽ trống</span>' : ''}<button class="b8sbtn" data-del="${esc(n)}" style="margin-left:auto" ${offUi.busy ? 'disabled' : ''}>Xóa</button></div>`; }).join('');
    }
    const add = $('#b8s-of-add'), name = $('#b8s-of-name'); if (add) add.disabled = offUi.busy; if (name) name.disabled = offUi.busy;
    const st = $('#b8s-of-st'); if (st && st.textContent !== offUi.msg) offSay(offUi.msg, offUi.bad);
  }
  async function officesAdd() {
    if (offUi.busy) return;
    const inp = $('#b8s-of-name'); if (!inp) return;
    offUi.busy = true; offSay('Đang kiểm tra tên bưu cục…'); renderOffices();
    try {
      const r = await officeLookup(inp.value);
      if (!r.ok) { offSay(r.msg, true); return; }
      // Chạy thử lấy đơn: bưu cục không có đơn nào hôm nay thì coi là bưu cục rỗng, không thêm
      offSay(`Tên đúng. Đang chạy thử lấy dữ liệu của “${r.name}”…`);
      state.sites[plain(r.name)] = { id: r.id, code: r.code };
      let probe;
      try { probe = await apiFetchOffice(r.name); }
      catch (e) { delete state.sites[plain(r.name)]; offSay(`Không lấy được dữ liệu thử của “${r.name}” (${e && e.message || e}). Chưa thêm.`, true); return; }
      const nOrders = Math.max(probe.items.length, Number(probe.total) || 0), nEmp = probe.t && probe.t.emp && probe.t.emp.rows ? probe.t.emp.rows.length : 0;
      if (nOrders === 0) { delete state.sites[plain(r.name)]; offSay(`Bưu cục “${r.name}” rỗng: ${nOrders} đơn, ${nEmp} nhân viên. Không thêm.`, true); return; }
      saveSites();
      OFFICES.push(r.name); offHidden.delete(r.name); offHidden.delete(plain(r.name)); saveOffices();
      inp.value = '';
      const ids = codIds();
      offSay(`Đã thêm “${r.name}” (${nOrders} đơn, ${nEmp} nhân viên).${Number.isFinite(ids[plain(r.name)]) ? '' : ' Nhập “Mã COD” ở dòng của bưu cục để có bảng COD.'} Bấm “Tải số liệu” ở tab ANM để lấy số.`);
      log(`Đã thêm bưu cục ${r.name}.`);
      if (state.anm.sel) { state.anm.sel.add(r.name); savePrefs(); }
      renderOfficeList();   // ô chọn bưu cục ở tab Tổng quan
      if (state.view === 'anm') renderAnm();
    } finally { offUi.busy = false; const b = $('#b8s-of-list'); if (b) delete b.dataset.sig; renderOffices(); }
  }
  function officesRemove(name) {
    const i = OFFICES.indexOf(name);
    if (i < 0) return;
    if (!window.confirm(`Xóa “${name}” khỏi danh sách theo dõi?\nLịch sử đã lưu không bị xóa, chỉ không còn hiện.`)) return;
    OFFICES.splice(i, 1); offHidden.add(name); saveOffices();
    if (state.anm.sel) { state.anm.sel.delete(name); savePrefs(); }
    offSay(`Đã xóa “${name}”.`); const b = $('#b8s-of-list'); if (b) delete b.dataset.sig; renderOffices(); renderOfficeList();
    if (state.view === 'anm') renderAnm();
  }
  function officesSaveCod(name, val) {
    const v = parseInt(val, 10), k = plain(name);
    if (!(v > 0)) {   // xóa mã đã nhập tay (mã có sẵn trong script vẫn dùng)
      try { const o = JSON.parse(localStorage.getItem('b8s_codsites') || '{}') || {}; delete o[k]; localStorage.setItem('b8s_codsites', JSON.stringify(o)); } catch { /* bỏ qua */ }
    } else saveCodIds({ [k]: v });
    const b = $('#b8s-of-list'); if (b) delete b.dataset.sig; renderOffices();
  }

  // ----- Tự dò bưu cục: chạy khi bấm “Tải số liệu” (ANM) hoặc “Chuyển bưu cục” (Tổng quan) -----
  const DISC_MAX_NEW = 12, DISC_MAX_PROBE = 80;
  const disc = { on: true, busy: false, undo: null, pick: null, note: '', warn: false, empty: {}, unknownOwner: false };   // empty: bưu cục 0 đơn đã kiểm tra trong ngày (không dò lại)
  function loadDisc() {
    try { const o = JSON.parse(localStorage.getItem('b8s_disc') || '{}') || {}; if (typeof o.on === 'boolean') disc.on = o.on; } catch { /* bỏ qua */ }
    try { const n = Number(localStorage.getItem('b8s_owner')); if (Number.isFinite(n) && n > 0) state.ownerSiteId = n; } catch { /* bỏ qua */ }
  }
  function saveDisc() { try { localStorage.setItem('b8s_disc', JSON.stringify({ on: disc.on })); } catch { /* bỏ qua */ } }
  // Lấy toàn bộ danh sách bưu cục tài khoản này được xem. Thử kèm ownerSiteId (nếu đã biết) rồi thử không kèm.
  async function discList() {
    const owners = []; if (Number.isFinite(state.ownerSiteId)) owners.push(state.ownerSiteId); owners.push(null);
    let lastErr = null;
    for (const ow of owners) {
      try {
        const all = []; let total = Infinity;
        for (let pg = 1; pg <= 20 && all.length < total; pg++) {
          const body = { pageNumber: pg, pageSize: 200, fuzzyNameOrCode: '', limitSiteByAuth: true, statusList: ['ENABLE', 'SHUTDOWN'] };
          if (ow !== null) body.ownerSiteId = ow;
          const d = await apiPost(API_SITE, body), arr = asArray(d);
          const t = Number(d && d.total); total = Number.isFinite(t) && t >= 0 ? t : arr.length;
          if (!arr.length) break;
          all.push(...arr);
        }
        if (all.length) return all;
        lastErr = new Error('hệ thống trả danh sách trống');
      } catch (e) { lastErr = e; }
    }
    throw lastErr || new Error('không lấy được danh sách');
  }
  // Số đơn hôm nay của một bưu cục (gọi nhỏ, chỉ lấy tổng)
  async function discProbe(c) {
    const d = await apiPost(API_DET, { isContainsSecSite: true, siteId: c.id, siteCode: c.code, shouldDispType: '1', searchType: '1', pageSize: 20, currentPage: 1 });
    return Math.max(asArray(d).length, Number(d && d.total) || 0);
  }
  function discRender() {
    const box = $('#b8s-disc'); if (!box) return;
    let h = '';
    if (disc.pick) {
      const warnOwner = disc.unknownOwner ? `<div style="margin-bottom:6px"><b>⚠️ Chưa xác định được chính xác bưu cục bạn quản lý.</b> Danh sách dưới đây là mọi bưu cục có đơn hôm nay mà tài khoản đang đăng nhập được phép xem, có thể gồm bưu cục KHÔNG phải của bạn. Hãy vào ô “Bưu cục” trên trang web, gõ/chọn 1 bưu cục của bạn một lần, rồi quay lại bấm “Tải số liệu” để tool tự nhận đúng — hoặc tự tích đúng bưu cục của bạn bên dưới, đừng bấm “Tích hết” nếu không chắc.</div>` : '';
      h = `<div class="al a">${warnOwner}<b>Tìm thấy ${disc.pick.length} bưu cục có đơn.</b> ${disc.unknownOwner ? 'Tool chưa rõ bưu cục của bạn nên hiện hết ra đây.' : `Nhiều hơn ${DISC_MAX_NEW} nên tool không tự thêm.`} Tích những bưu cục bạn muốn theo dõi:<div style="max-height:150px;overflow:auto;margin:4px 0">${disc.pick.map((c, i) => `<label style="display:inline-block;min-width:230px"><input type="checkbox" data-dpick="${i}" ${c.on === false ? '' : 'checked'}> ${esc(c.name)} <span class="muted">(${c.n} đơn)</span></label>`).join('')}</div><button class="b8sbtn solid" data-d="padd">Thêm các bưu cục đã tích</button> <button class="b8sbtn" data-d="pall">Tích hết</button> <button class="b8sbtn" data-d="pnone">Bỏ tích hết</button> <button class="b8sbtn" data-d="pskip">Bỏ qua</button></div>`;
    } else if (disc.note) {
      h = `<div class="al ${disc.warn ? 'a' : 'ok'}" style="display:flex;gap:8px;align-items:center"><span style="flex:1">${esc(disc.note)}</span>${disc.undo && disc.undo.length ? '<button class="b8sbtn" data-d="undo" title="Bỏ các bưu cục vừa được tự thêm">Hoàn tác</button>' : ''}<button class="b8sbtn ghost" data-d="close">×</button></div>`;
    }
    box.style.display = h ? 'block' : 'none';
    if (box.dataset.h !== h) { box.dataset.h = h; box.innerHTML = h; }
  }
  function discAdd(cands) {
    const names = [];
    for (const c of cands) {
      if (OFFICES.includes(c.name)) continue;
      state.sites[plain(c.name)] = { id: c.id, code: c.code };
      OFFICES.push(c.name); offHidden.delete(c.name); offHidden.delete(plain(c.name)); names.push(c.name);
      if (state.anm.sel) state.anm.sel.add(c.name);
    }
    if (names.length) { saveSites(); saveOffices(); savePrefs(); }
    if (names.length) { const b = $('#b8s-of-list'); if (b) delete b.dataset.sig; renderOffices(); renderOfficeList(); }
    return names;
  }
  function discUndo() {
    const names = (disc.undo || []).slice();
    for (const n of names) { const i = OFFICES.indexOf(n); if (i >= 0) OFFICES.splice(i, 1); offHidden.add(n); if (state.anm.sel) state.anm.sel.delete(n); }
    if (names.length) { saveOffices(); savePrefs(); }
    disc.undo = null; disc.note = names.length ? `Đã hoàn tác: bỏ ${names.join(', ')} khỏi danh sách theo dõi.` : ''; disc.warn = false;
    log(disc.note);
    const b = $('#b8s-of-list'); if (b) delete b.dataset.sig;
    renderOffices(); renderOfficeList(); discRender(); if (state.view === 'anm') renderAnm();
  }
  function discPickDone(addIt) {
    const p = disc.pick || []; disc.pick = null;
    const chosen = addIt ? p.filter(c => c.on !== false) : [];
    const names = discAdd(chosen);
    disc.undo = names.length ? names : null;
    disc.note = names.length ? `Đã thêm ${names.length} bưu cục: ${names.join(', ')}. Bấm “Tải số liệu” lại để lấy số.` : 'Không thêm bưu cục nào.'; disc.warn = !names.length;
    log(disc.note); discRender(); if (state.view === 'anm') renderAnm();
  }
  // Trả về {added, ...} hoặc null nếu không chạy. Không bao giờ ném lỗi: lỗi chỉ được báo, danh sách cũ giữ nguyên.
  async function discoverOffices() {
    if (!disc.on || disc.busy) return null;
    disc.busy = true; disc.pick = null;
    const st = (m) => { log(m); };
    try {
      st('Đang dò các bưu cục bạn quản lý…');
      let list;
      try { list = await discList(); }
      catch (e) {
        const m = `Không dò được danh sách bưu cục (${e && e.message || e}). Giữ nguyên danh sách hiện tại${OFFICES.length ? '' : '; bạn có thể thêm bưu cục thủ công ở Cài đặt'}.`;
        disc.note = m; disc.warn = true; disc.undo = null; log(m, true); discRender(); return null;
      }
      const pk = discPick(list, new Set(OFFICES.map(plainOf)), offHidden);
      // Bưu cục đang lưu nhưng tài khoản hiện đăng nhập không còn thấy trong danh sách được phép xem hôm nay: không còn thuộc tài khoản này, tự loại khỏi danh sách theo dõi
      const staleRemoved = OFFICES.filter(o => !pk.allNames.has(plainOf(o)));
      if (staleRemoved.length) {
        for (const n of staleRemoved) { const i = OFFICES.indexOf(n); if (i >= 0) OFFICES.splice(i, 1); delete state.sites[plain(n)]; }
        saveOffices(); saveSites();
        const b = $('#b8s-of-list'); if (b) delete b.dataset.sig;
      }
      const today = dateKey(new Date());
      const todo = pk.cand.filter(c => disc.empty[c.id] !== today);
      const left = Math.max(0, todo.length - DISC_MAX_PROBE), batch = todo.slice(0, DISC_MAX_PROBE);
      const ok = [], empty = pk.cand.filter(c => disc.empty[c.id] === today).map(c => c.name), failed = [];
      let p = 0;
      const worker = async () => {
        while (p < batch.length) {
          const c = batch[p++];
          try { const n = await discProbe(c); if (n > 0) ok.push(Object.assign({}, c, { n })); else { empty.push(c.name); disc.empty[c.id] = today; } }
          catch { failed.push(c.name); }
        }
      };
      await Promise.all([worker(), worker(), worker()]);
      ok.sort((a, b) => a.name.localeCompare(b.name, 'vi'));
      const unknownOwner = !OFFICES.length && !Number.isFinite(state.ownerSiteId);   // lần đầu cài, chưa rõ bưu cục của tài khoản: luôn hỏi lại, không tự thêm
      let added = [];
      if (ok.length > DISC_MAX_NEW || unknownOwner) {
        disc.pick = unknownOwner ? ok.map(c => Object.assign({}, c, { on: false })) : ok;
        disc.note = ''; disc.undo = null; disc.unknownOwner = unknownOwner;
      }
      else { disc.unknownOwner = false; added = discAdd(ok); disc.undo = added.length ? added : null; }
      const quiet = !added.length && !empty.length && !failed.length && !left && !disc.pick && !staleRemoved.length && OFFICES.length;
      const sum = discSummary(pk, added, empty, failed, left, staleRemoved);
      if (!disc.pick) {
        disc.note = quiet ? '' : (!OFFICES.length ? `${sum} Chưa có bưu cục nào có đơn hôm nay.` : sum);
        disc.warn = !added.length || !!staleRemoved.length;
      }
      if (!quiet) log(sum, !!failed.length || !!staleRemoved.length);
      discRender();
      if (added.length || staleRemoved.length) { renderOffices(); renderOfficeList(); if (state.view === 'anm') renderAnm(); }
      return { added, empty, failed, staleRemoved, picked: !!disc.pick };
    } catch (e) {
      log(`Dò bưu cục lỗi: ${e && e.message || e}`, true); return null;
    } finally { disc.busy = false; }
  }
  // Nút "Tải danh sách bưu cục" (tab Tổng quan): hỏi thẳng hệ thống xem tài khoản đang đăng nhập quản lý đúng những bưu cục nào,
  // rồi THAY THẾ HOÀN TOÀN danh sách đang theo dõi bằng danh sách đó (không gộp/giữ bưu cục cũ). Danh sách này dùng chung cho cả tab ANM.
  async function refreshOfficeList() {
    if (state.busy || disc.busy || state.ofList.busy) return;
    state.ofList.busy = true; renderDashboard();
    try {
      // bưu cục đang hiện ra (kể cả bưu cục chỉ còn trong lịch sử cũ, không nằm trong OFFICES) — để biết cái nào cần ẩn đi.
      // Tự tính thay vì gọi anmOffices() vì anmOffices() thuộc phần tab ANM có thể đang bị khóa (bản chia sẻ) chưa nhập mật khẩu.
      const before = new Set(OFFICES);
      for (const r of state.hist.values()) { const o = canonOffice(r.office, OFFICES); if (!offHidden.has(o)) before.add(o); }
      const list = await discList();   // toàn bộ bưu cục tài khoản hiện đăng nhập được phép xem (ưu tiên lọc theo ownerSiteId nếu đã biết)
      const pk = discPick(list, new Set(), new Set());   // known/hidden để trống: lấy hết bưu cục hợp lệ, không phân biệt mới/cũ
      // Chỉ tính là "đang quản lý" khi hôm nay có đơn — loại bưu cục rỗng (0 đơn), tránh lấy nhầm bưu cục không hoạt động
      const todayK = dateKey(new Date()), batch = pk.cand.slice(0, DISC_MAX_PROBE);
      const okC = [], emptyC = [];
      let pi = 0;
      const probeWorker = async () => {
        while (pi < batch.length) {
          const c = batch[pi++];
          try { const n = await discProbe(c); if (n > 0) okC.push(c); else emptyC.push(c.name); }
          catch { /* lỗi kiểm tra: bỏ qua, không tính vào danh sách */ }
        }
      };
      await Promise.all([probeWorker(), probeWorker(), probeWorker()]);
      const skippedLeft = pk.cand.length - batch.length;
      const names = okC.map(c => c.name).sort((a, b) => a.localeCompare(b, 'vi'));
      const nameSet = new Set(names);
      OFFICES.splice(0, OFFICES.length, ...names);
      // Thay thế hoàn toàn: bưu cục nào trước đó có hiện ra (kể cả chỉ do còn bản lưu cũ trong lịch sử) mà giờ không còn
      // trong danh sách mới thì ẩn hẳn đi, để không bị lịch sử cũ kéo lại hiện ra lần nữa.
      for (const n of before) if (!nameSet.has(n)) offHidden.add(n);
      for (const n of names) { offHidden.delete(n); offHidden.delete(plain(n)); }
      for (const k of Object.keys(state.sites)) delete state.sites[k];
      for (const c of okC) state.sites[plain(c.name)] = { id: c.id, code: String(c.code) };
      state.anm.sel = null;   // chọn lại tất cả theo danh sách mới
      saveOffices(); saveSites(); savePrefs();
      const b = $('#b8s-of-list'); if (b) delete b.dataset.sig;
      state.ofList.at = Date.now();
      const extra = [];
      if (emptyC.length) extra.push(`bỏ qua ${emptyC.length} bưu cục rỗng hôm nay (${emptyC.join(', ')})`);
      if (skippedLeft > 0) extra.push(`còn ${skippedLeft} bưu cục chưa kiểm tra kịp, bấm tải lại để kiểm tra tiếp`);
      state.ofList.note = (names.length ? `Đã tải ${names.length} bưu cục đang quản lý (có đơn hôm nay)` : 'Không có bưu cục nào có đơn hôm nay') + (extra.length ? `; ${extra.join('; ')}` : '') + '.';
      state.ofList.warn = !names.length || !!emptyC.length || skippedLeft > 0;
      log(state.ofList.note, !names.length);
    } catch (e) {
      state.ofList.note = `Không tải được danh sách bưu cục (${e && e.message || e}). Danh sách cũ vẫn giữ nguyên.`;
      state.ofList.warn = true;
      log(state.ofList.note, true);
    } finally {
      state.ofList.busy = false;
      renderOffices(); renderOfficeList(); renderDashboard(); if (state.view === 'anm') renderAnm();
    }
  }
  async function resolveSite(name) {
    const key = plain(name);
    if (state.sites[key]) return state.sites[key];
    const nm = x => { const v = x.nameLocale; if (typeof v === 'string') return v; if (v && typeof v === 'object') return Object.values(v).find(t => typeof t === 'string') || ''; return x.name || ''; };
    let hit = null, n = 0;
    for (const auth of [true, false]) {
      const list = asArray(await apiPost(API_SITE, { pageNumber: 1, pageSize: 200, fuzzyNameOrCode: name, limitSiteByAuth: auth, statusList: ['ENABLE', 'SHUTDOWN'] }));
      n = list.length;
      hit = list.find(x => plain(nm(x)) === key && x.status !== 'SHUTDOWN') || list.find(x => plain(nm(x)) === key);
      if (hit) break;
    }
    if (!hit || !hit.id || !hit.code) throw new Error(`không tìm thấy mã bưu cục “${name}” (API trả ${n} kết quả)`);
    state.sites[key] = { id: hit.id, code: String(hit.code) };
    saveSites();
    return state.sites[key];
  }
  // Gọi API lấy toàn bộ đơn + nhân viên của một bưu cục và dựng thành bảng giống trên trang
  async function apiFetchOffice(name) {
    const site = await resolveSite(name);
    const base = { isContainsSecSite: true, siteId: site.id, siteCode: site.code, shouldDispType: '1' };
    const det = Object.assign({}, base, { searchType: '1' });
    const sizes = state.apiPs ? [state.apiPs] : [1000, 500, 200, 100, 50];
    const empP = apiPost(API_EMP, base);
    empP.catch(() => { /* lỗi sẽ được báo khi đợi kết quả bên dưới; tránh báo lỗi chưa xử lý nếu API đơn lỗi trước */ });
    let first = null, lastErr = null;
    for (const ps of sizes) {
      try { first = await apiPost(API_DET, Object.assign({}, det, { pageSize: ps, currentPage: 1 })); state.apiPs = ps; break; }
      catch (e) { lastErr = e; if (state.apiPs) state.apiPs = 0; }
    }
    if (!first) { if (state.apiPs === 0) { state.apiPs = null; } throw lastErr || new Error('API đơn không trả dữ liệu'); }
    const empData = await empP;
    let items = ((first && first.list) || []).slice();
    const total = Number(first && first.total) || 0, per = items.length;
    let calls = 1, dropped = 0;
    if (per && items.length < total) {
      const pages = Math.min(400, Math.ceil(total / per));
      const nums = []; for (let p = 2; p <= pages; p++) nums.push(p);
      for (let i = 0; i < nums.length; i += 6) {
        const got = await Promise.all(nums.slice(i, i + 6).map(p => apiPost(API_DET, Object.assign({}, det, { pageSize: per, currentPage: p })).then(d => (d && d.list) || [])));
        got.forEach(g => items.push(...g)); calls += got.length;
      }
      // Nhiều trang: dữ liệu đổi giữa các lần gọi có thể làm một dòng xuất hiện ở hai trang; chỉ loại các dòng GIỐNG HỆT nhau hoàn toàn
      const sg = new Set(), uniq = [];
      for (const it of items) { let k; try { k = JSON.stringify(it); } catch { k = null; } if (k !== null) { if (sg.has(k)) { dropped++; continue; } sg.add(k); } uniq.push(it); }
      items = uniq;
    }
    const t = apiToTables(items, asArray(empData));
    const diag = { rows: items.length, pages: calls, dup: t.dupN, empty: t.emptyN, dropped };
    state.sign = { keys: t.signKeys, okN: t.okN, signN: t.signN };
    if (t.okN > 0 && t.signN === 0 && !state.signWarned) { state.signWarned = true; log(`${name}: không thấy trường giờ ký nhận trong đơn nên chưa có “GTC đầu tiên”. Bấm “Chẩn đoán API” rồi gửi dòng signTimeKeys.`, true); }
    if (t.unknown.length) log(`${name}: có mã trạng thái chưa biết (${t.unknown.join(', ')}), tạm xếp vào “khác”.`, true);
    return { t, items, total, diag };
  }
  async function apiLoadOffice(name) {
    const { t, items, total, diag } = await apiFetchOffice(name);
    return apiSave(name, t, items, total, diag);
  }
  async function apiSave(name, t, items, total, diag) {
    const now = new Date();
    const calc = computeDashboard({ orders: t.orders, emp: t.emp, tiktok: parseList(state.tiktok), statusMap: state.statusMap, now, rules: state.dm, skip: shipSkipFn(name) });
    // v3.35: Tỷ lệ ODR giờ tính thẳng từ cột “Thời hạn giao hàng của đối tác” trong computeDashboard (calc.extra.odr2) — không gọi Kanban nữa.
    const d = decideSave({ hasOrders: true, pageOffice: name, startOffice: name, expectedOffice: name, ovNeed: total, computedTotal: calc.cards.all.total, diag });
    const rec = buildRecord({ office: canonOffice(name, OFFICES) || name, now, calc, verified: !!d.verified });
    state.kb.sg.set(rec.office, { date: rec.date, at: Date.parse(rec.savedAt) || Date.now(), sg: signMinsOf(t.orders, state.statusMap, new Date()) });
    if (!d.save) { return { saved: false, reason: d.reason, rec: d.force ? rec : null }; }
    await putRecords([rec]);
    state.lastSaved = new Date();
    log(`Đã lưu lịch sử ${rec.office} ngày ${fmtDMY(rec.date)} (API, ${fmt(items.length)} đơn).`);
    return { saved: true, verified: !!d.verified };
  }

  async function trySaveHistory(expect, startOffice) {
    state.pending = null;
    const c = getCalc();
    // Nếu không đọc được tên bưu cục từ trang thì dùng đúng bưu cục tool vừa chọn (tool đã chọn và chờ trang tải xong)
    const pageOffice = readOfficeInfo().office || expect || state.office || '';
    const m = (document.body?.innerText || '').match(/Tổng đơn cần phát hàng\s*([\d,.]+)/i);
    const d = decideSave({
      hasOrders: !!(c && state.orders), pageOffice, startOffice, expectedOffice: expect,
      ovNeed: m ? parseNum(m[1]) : (state.ov && Number.isFinite(state.ov.need) ? state.ov.need : NaN), computedTotal: c ? c.cards.all.total : 0
    });
    if (!c) return { saved: false, reason: 'Chưa tính được số liệu từ bảng đơn.' };
    const now = new Date();
    // v3.35: Tỷ lệ ODR giờ tính thẳng từ cột “Thời hạn giao hàng của đối tác” trong computeDashboard (c.extra.odr2) — không gọi Kanban nữa.
    const rec = d.reason && !d.force ? null : buildRecord({ office: canonOffice(pageOffice, OFFICES), now, calc: c, verified: !!d.verified });
    if (d.save && rec) {
      try {
        await putRecords([rec]);
        state.lastSaved = new Date();
        log(`Đã lưu lịch sử ${rec.office} ngày ${fmtDMY(rec.date)}${d.verified ? '' : ' (chưa đối chiếu được tổng đơn với trang)'}.`);
        return { saved: true, verified: !!d.verified };
      } catch (e) { log(`Không lưu được lịch sử: ${e.message}`, true); return { saved: false, reason: `Lỗi ghi kho lưu trữ: ${e.message}` }; }
    }
    log(`Chưa lưu lịch sử: ${d.reason}`, true);
    if (d.force && rec) state.pending = { reason: d.reason, rec };
    return { saved: false, reason: d.reason, rec: d.force ? rec : null };
  }

  async function forceSave() {
    const p = state.pending;
    if (!p) return;
    try {
      await putRecords([{ ...p.rec, verified: false }]);
      state.lastSaved = new Date(); state.pending = null;
      log(`Đã lưu lịch sử ${p.rec.office} ngày ${fmtDMY(p.rec.date)} (lưu thủ công).`);
    } catch (e) { log(`Không lưu được lịch sử: ${e.message}`, true); }
    renderUI();
  }

  function backupInfo() {
    let last = 0; try { last = Number(localStorage.getItem('b8s_lastBackup') || 0); } catch { /* bỏ qua */ }
    return { last, due: backupDue(state.hist.size, last, Date.now()) };
  }

  function downloadText(name, text, type) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([text], { type }));
    a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  function backupHistory() {
    const records = [...state.hist.values()];
    if (!records.length) { log('Chưa có lịch sử để sao lưu.', true); return; }
    downloadText(`800BEST_LichSu_${dateKey(new Date())}.json`, JSON.stringify({ app: '800best-history', version: 1, exportedAt: new Date().toISOString(), records, snaps: state.snap.list }), 'application/json');
    try { localStorage.setItem('b8s_lastBackup', String(Date.now())); } catch { /* bỏ qua */ }
    log(`Đã tạo file sao lưu ${records.length} bản ghi.`);
    renderUI();
  }
  async function restoreHistory(file) {
    let obj;
    try { obj = JSON.parse(await file.text()); } catch { log('File không đọc được (không phải JSON).', true); return; }
    const v = validateBackup(obj);
    if (!v.ok) { log(v.error, true); return; }
    // bản chụp theo mốc giờ (nếu file có): chỉ thêm những bản chưa có hoặc cũ hơn
    const snapsIn = (Array.isArray(obj.snaps) ? obj.snaps : []).map(normalizeSnap).filter(Boolean);
    const snapPut = snapsIn.filter(r => { const cur = snapAt(state.snap.idx, r.office, r.date, r.slot); return !cur || String(r.runAt) > String(cur.runAt); });
    if (snapPut.length && snapStore) {
      try { await snapStore.putMany(snapPut); await loadSnaps(); log(`Đã khôi phục ${snapPut.length} bản chụp theo mốc giờ.`); }
      catch (e) { log(`Không khôi phục được bản chụp theo mốc giờ: ${e.message}`, true); }
    }
    const plan = planMerge(state.hist, v.records);
    if (!plan.put.length) { log(`Không có gì cần khôi phục (${plan.kept} bản ghi đã có bản mới hơn hoặc bằng).`); return; }
    if (!confirm(`Khôi phục lịch sử?\nThêm mới: ${plan.added} ngày\nCập nhật (bản trong file mới hơn): ${plan.updated} ngày\nGiữ nguyên (đã có bản mới hơn): ${plan.kept} ngày`)) return;
    try { await putRecords(plan.put); try { localStorage.setItem('b8s_lastBackup', String(Date.now())); } catch { /* bỏ qua */ } log(`Đã khôi phục: thêm ${plan.added}, cập nhật ${plan.updated}, giữ nguyên ${plan.kept}.`); }
    catch (e) { log(`Khôi phục thất bại: ${e.message}`, true); }
    renderUI();
  }

  /* ---- Tab Lịch sử ---- */
  const HCOLS = [
    { k: 'total', label: 'Tổng đơn', kind: 'avg' },
    { k: 'scanRate', label: 'Tỷ lệ quét giao', kind: 'pct' },
    { k: 'pod', label: 'POD2400', kind: 'pct', n: 'podN' },
    { k: 'podTT', label: 'POD TikTok', kind: 'pct', n: 'podTTN' },
    { k: 'gtcTrip', label: 'GTC chuyến', kind: 'pct' },
    { k: 'gtcTotal', label: 'GTC Total', kind: 'pct' },
    { k: 'backlog', label: 'Backlog 3 ngày', kind: 'avg', last: 'backlogLast' },
    { k: 'odr', label: 'TikTok đến hạn ODR', kind: 'avg', last: 'odrLast' },
    { k: 'nWork', label: 'Shipper đi làm', kind: 'avg' }
  ];

  const histOffices = () => {
    const set = new Set(OFFICES);
    for (const r of state.hist.values()) { const o = canonOffice(r.office, OFFICES); if (!offHidden.has(o)) set.add(o); }
    return [...set];
  };
  function histFilter() {
    const hv = state.hv, today = dateKey(new Date());
    const office = hv.office || (state.office ? canonOffice(state.office, OFFICES) : '*');
    const rg = hv.preset === 'custom' ? { from: hv.from || '0000-01-01', to: hv.to || '9999-12-31' } : rangeFor(hv.preset, today);
    const recs = [...state.hist.values()]
      .filter(r => (office === '*' || sameOffice(r.office, office)) && r.date >= rg.from && r.date <= rg.to)
      .sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : a.office < b.office ? -1 : 1);
    return { office, rg, recs };
  }
  function trendRows(recs, group) { return withDeltas(groupRecords(recs, group)); }

  function hcell(col, m, delta, multi) {
    const v = m[col.k];
    if (v === null || v === undefined || !Number.isFinite(v)) return '<td>—</td>';
    if (col.kind === 'pct') {
      const n = col.n && m[col.n], tip = n ? `${fmt(n[0])} / ${fmt(n[1])} đơn` : '';
      const d = delta && Number.isFinite(delta[col.k]) ? delta[col.k] : null;
      const dd = d === null || Math.abs(d) < 0.05 ? '' : `<small class="${d > 0 ? 'up' : 'dn'}">${d > 0 ? '▲' : '▼'}${round1(Math.abs(d)).toLocaleString('vi-VN')}</small>`;
      return `<td class="${rateClass(v)}" title="${esc(tip)}">${esc(fmtVal(v, 'pct'))}${dd}</td>`;
    }
    const last = col.last && multi && m[col.last] !== null && m[col.last] !== undefined ? `Ngày cuối kỳ: ${fmt(m[col.last])}` : '';
    return `<td title="${esc(last)}">${esc(fmtAvg(v))}</td>`;
  }

  function renderHistory() {
    const box = $('#b8s-hist');
    if (!box) return;
    const hv = state.hv, f = histFilter(), today = dateKey(new Date());
    const offs = histOffices();
    const bi = backupInfo();
    const opt = (v, cur, l) => `<option value="${esc(v)}" ${v === cur ? 'selected' : ''}>${esc(l)}</option>`;
    const cmpDate = hv.cmpDate || ([...state.hist.values()].map(r => r.date).sort().pop() || today);
    const controls = `
      <div class="b8s-panel">
        <div class="row">
          <label>Bưu cục<br><select data-h="office">${opt('*', f.office, 'Tất cả (gộp)')}${offs.map(o => opt(o, f.office, o)).join('')}</select></label>
          <label>Xem<br><select data-h="mode">${opt('trend', hv.mode, 'Xu hướng theo kỳ')}${opt('ship', hv.mode, 'Shipper (tổng hợp)')}${opt('cmp', hv.mode, 'So sánh các bưu cục')}</select></label>
          ${hv.mode === 'cmp' ? `<label>Ngày<br><input type="date" data-h="cmpDate" value="${esc(cmpDate)}"></label>` : `
          <label>Khoảng thời gian<br><select data-h="preset">${[['d7', '7 ngày gần nhất'], ['d30', '30 ngày gần nhất'], ['wk', 'Tuần này'], ['lastwk', 'Tuần trước'], ['mo', 'Tháng này'], ['lastmo', 'Tháng trước'], ['all', 'Tất cả'], ['custom', 'Tự chọn…']].map(([v, l]) => opt(v, hv.preset, l)).join('')}</select></label>
          ${hv.preset === 'custom' ? `<label>Từ<br><input type="date" data-h="from" value="${esc(hv.from)}"></label><label>Đến<br><input type="date" data-h="to" value="${esc(hv.to)}"></label>` : ''}
          ${hv.mode === 'trend' ? `<label>Gộp theo<br><select data-h="group">${[['day', 'Ngày'], ['week', 'Tuần'], ['month', 'Tháng']].map(([v, l]) => opt(v, hv.group, l)).join('')}</select></label>` : ''}`}
          <span style="flex:1"></span>
          <button class="b8sbtn solid" data-h-act="backup">Sao lưu</button>
          <button class="b8sbtn solid" data-h-act="restore">Khôi phục</button>
          <input type="file" id="b8s-restore-file" accept=".json,application/json" style="display:none">
        </div>
        ${bi.due ? '<div class="b8s-warn" style="margin:6px 0 0">⚠ Đã hơn 7 ngày chưa sao lưu lịch sử. Bấm “Sao lưu” để tải file về máy (phòng khi xóa dữ liệu trình duyệt).</div>' : ''}
        ${state.histErr ? `<div class="b8s-warn" style="margin:6px 0 0">${esc(state.histErr)}</div>` : ''}
      </div>`;

    let body = '';
    if (!state.hist.size) {
      body = `<div class="b8s-empty"><b>Chưa có lịch sử</b><br>Mỗi lần bạn chọn bưu cục và tải xong số liệu, chỉ số của ngày hôm đó được tự lưu (chỉ số tổng và từng shipper, không lưu mã vận đơn hay thông tin người nhận). Tải lại trong ngày thì bản lưu của ngày đó được cập nhật.</div>`;
    } else if (hv.mode === 'cmp') {
      const recs = [...state.hist.values()].filter(r => r.date === cmpDate).sort((a, b) => a.office < b.office ? -1 : 1);
      if (!recs.length) body = `<div class="b8s-empty">Không có bản ghi nào của ngày ${esc(fmtDMY(cmpDate))}.</div>`;
      else {
        const rows = recs.map(r => ({ r, m: metricsOf(aggregate([r])) }));
        const tot = metricsOf(aggregate(recs));
        body = `<div class="b8s-tw"><table><thead><tr><th class="l g-nav">Bưu cục</th>${HCOLS.map(c => `<th class="g-gh">${esc(c.label)}</th>`).join('')}</tr></thead><tbody>${
          rows.map(({ r, m }) => `<tr><td class="l"><b>${esc(r.office)}</b>${r.verified ? '' : ' <span title="Chưa đối chiếu được tổng đơn với trang">⚠</span>'}</td>${HCOLS.map(c => hcell(c, m, null, false)).join('')}</tr>`).join('')
        }${recs.length > 1 ? `<tr class="total"><td class="l">TỔNG</td>${HCOLS.map(c => hcell(c, tot, null, false)).join('')}</tr>` : ''}</tbody></table></div>`;
      }
    } else if (hv.mode === 'ship') {
      let list = aggregateShippers(f.recs);
      const sk = hv.ssort.key, dir = hv.ssort.dir;
      list.sort((a, b) => {
        const x = a[sk], y = b[sk];
        if (typeof x === 'string' || typeof y === 'string') return String(x ?? '').localeCompare(String(y ?? ''), 'vi') * dir;
        const nx = Number.isFinite(x) ? x : -Infinity, ny = Number.isFinite(y) ? y : -Infinity;
        return (nx - ny) * dir;
      });
      const cols = [['name', 'Shipper', 'text'], ['office', 'Bưu cục', 'text'], ['days', 'Số ngày đi làm', 'int'], ['sc', 'Đã quét giao', 'int'], ['avgSc', 'TB/ngày', 'avg'], ['ok', 'Giao thành công', 'int'], ['rateTrip', '% GTC chuyến', 'pct'], ['tt', 'Đơn TikTok', 'int'], ['tk', 'TikTok giao thành công', 'int'], ['rateTT', '% GTC TikTok', 'pct'], ['sg', 'Đã ký nhận (tab NV)', 'int']];
      body = list.length ? `<div class="b8s-tw"><table><thead><tr><th class="g-nav">#</th>${cols.map(([k, l]) => `<th class="g-gh ${k === 'name' ? 'l' : ''}" data-sk="${k}">${esc(l)}${hv.ssort.key === k ? (dir > 0 ? ' ▲' : ' ▼') : ''}</th>`).join('')}</tr></thead><tbody>${
        list.map((x, i) => `<tr><td>${i + 1}</td>${cols.map(([k, , kind]) => kind === 'text' ? `<td class="l">${esc(x[k])}</td>` : kind === 'pct' ? `<td class="${rateClass(x[k])}">${esc(fmtVal(x[k], 'pct'))}</td>` : kind === 'avg' ? `<td>${x[k] === null ? '—' : esc(fmtAvg(x[k]))}</td>` : `<td>${x[k] === null ? '—' : esc(fmt(x[k]))}</td>`).join('')}</tr>`).join('')
      }</tbody></table></div>` : '<div class="b8s-empty">Không có dữ liệu shipper trong khoảng này.</div>';
    } else {
      if (!f.recs.length) body = '<div class="b8s-empty">Không có bản ghi nào trong khoảng thời gian này.</div>';
      else {
        const list = trendRows(f.recs, hv.group), show = [...list].reverse();
        const whole = list.length > 1 ? metricsOf(aggregate(f.recs)) : null;
        const canDel = hv.group === 'day' && f.office !== '*';
        body = `<div class="b8s-tw"><table><thead><tr><th class="l g-nav">Kỳ</th><th class="g-nav">Số ngày</th>${HCOLS.map(c => `<th class="g-gh">${esc(c.label)}</th>`).join('')}${canDel ? '<th class="g-nav"></th>' : ''}</tr></thead><tbody>${
          (whole ? `<tr class="total"><td class="l">Cả kỳ</td><td>${fmt(whole.days)}</td>${HCOLS.map(c => hcell(c, whole, null, true)).join('')}${canDel ? '<td></td>' : ''}</tr>` : '') +
          show.map(g => `<tr><td class="l"><b>${esc(g.label)}</b>${g.recs.some(r => !r.verified) ? ' <span title="Có bản ghi chưa đối chiếu được tổng đơn với trang">⚠</span>' : ''}</td><td>${fmt(g.m.days)}</td>${HCOLS.map(c => hcell(c, g.m, g.delta, g.m.days > 1)).join('')}${canDel ? `<td><button class="b8sx" title="Xóa bản ghi ngày này" data-del="${esc(g.recs[0].id)}">✕</button></td>` : ''}</tr>`).join('')
        }</tbody></table></div>
        <div class="b8s-notes">Tỷ lệ của tuần/tháng được tính lại từ tổng số đơn của cả kỳ (không lấy trung bình các %). Tổng đơn, Backlog, TikTok đến hạn ODR, Shipper đi làm là trung bình mỗi ngày; rê chuột vào ô để xem số đơn hoặc giá trị ngày cuối kỳ. ▲▼ là chênh lệch điểm % so với kỳ trước.</div>`;
      }
    }
    box.innerHTML = controls + body;
  }

  function onHistChange(e) {
    const t = e.target, hv = state.hv;
    if (t.id === 'b8s-restore-file') { const f = t.files && t.files[0]; t.value = ''; if (f) restoreHistory(f); return; }
    const k = t.dataset && t.dataset.h;
    if (!k) return;
    hv[k] = t.value;
    if (k === 'preset' && t.value === 'custom' && !hv.from) { const r = rangeFor('d30', dateKey(new Date())); hv.from = r.from; hv.to = r.to; }
    renderHistory();
  }
  async function onHistClick(e) {
    const b = e.target.closest('[data-h-act]'), th = e.target.closest('th[data-sk]'), del = e.target.closest('[data-del]');
    if (b) {
      if (b.dataset.hAct === 'backup') backupHistory();
      else if (b.dataset.hAct === 'restore') { const fi = $('#b8s-restore-file'); if (fi) fi.click(); }
    } else if (th) {
      const key = th.dataset.sk, ss = state.hv.ssort;
      state.hv.ssort = ss.key === key ? { key, dir: -ss.dir } : { key, dir: (key === 'name' || key === 'office') ? 1 : -1 };
      renderHistory();
    } else if (del) {
      const id = del.dataset.del;
      if (!confirm(`Xóa bản ghi ${id.replace('|', ' – ')}?`)) return;
      try { await store.del(id); state.hist.delete(id); log(`Đã xóa bản ghi ${id}.`); } catch (er) { log(`Không xóa được: ${er.message}`, true); }
      renderUI();
    }
  }

  // Thêm các sheet lịch sử vào file Excel (theo bộ lọc đang chọn ở tab Lịch sử)
  function addHistorySheets(addSheet) {
    if (!state.hist.size) return;
    const f = histFilter();
    if (!f.recs.length) return;
    const pct = v => Number.isFinite(v) ? round1(v) : '';
    const avg = v => (v === null || v === undefined || !Number.isFinite(v)) ? '' : round1(v);
    const head = ['Kỳ', 'Số ngày', 'Tổng đơn (TB/ngày)', 'Tỷ lệ quét giao (%)', 'POD2400 (%)', 'POD TikTok (%)', 'GTC chuyến (%)', 'GTC Total (%)', 'Backlog 3 ngày (TB/ngày)', 'Backlog ngày cuối', 'TikTok đến hạn ODR (TB/ngày)', 'ODR ngày cuối', 'Shipper đi làm (TB/ngày)'];
    const line = (label, m) => [label, m.days, avg(m.total), pct(m.scanRate), pct(m.pod), pct(m.podTT), pct(m.gtcTrip), pct(m.gtcTotal), avg(m.backlog), avg(m.backlogLast), avg(m.odr), avg(m.odrLast), avg(m.nWork)];
    const list = trendRows(f.recs, state.hv.group);
    addSheet('Lich su', [head, ...list.map(g => line(g.label, g.m)), ...(list.length > 1 ? [line('CẢ KỲ', metricsOf(aggregate(f.recs)))] : [])]);
    addSheet('Theo ngay', [['Ngày', 'Bưu cục', 'Tổng đơn', 'Chưa quét giao', 'Giao thành công', 'Đơn mới', 'Đơn mới giao thành công', 'TikTok tổng', 'TikTok giao thành công', 'Backlog 3 ngày', 'Chờ duyệt', 'Giao thất bại', 'Nhập kho chưa quét', 'TikTok đến hạn ODR', 'Shipper đi làm', 'Đã đối chiếu'],
      ...f.recs.map(r => [r.date, r.office, r.all.total, r.all.unscanned, r.all.ok, r.all.newTotal, r.all.newOk, r.tt.total, r.tt.ok, r.backlog ? r.backlog.total : '', r.backlog ? r.backlog.wait : '', r.backlog ? r.backlog.fail : '', r.backlog ? r.backlog.unscanned : '', r.odrTT ?? '', r.nWork, r.verified ? 'Có' : 'Chưa'])]);
    const sh = aggregateShippers(f.recs);
    if (sh.length) {
      addSheet('Shipper tong', [['Shipper', 'Bưu cục', 'Số ngày đi làm', 'Đã quét giao', 'TB/ngày', 'Giao thành công', '% GTC chuyến', 'Đơn TikTok', 'TikTok giao thành công', '% GTC TikTok', 'Đã ký nhận (tab NV)'],
        ...sh.sort((a, b) => (b.rateTrip ?? -1) - (a.rateTrip ?? -1)).map(x => [x.name, x.office, x.days, x.sc, avg(x.avgSc), x.ok, pct(x.rateTrip), x.tt, x.tk, pct(x.rateTT), x.sg ?? ''])]);
      const day = [['Ngày', 'Bưu cục', 'Shipper', 'Đã quét giao', 'Giao thành công', '% GTC chuyến', 'Đơn TikTok', 'TikTok giao thành công', '% GTC TikTok', 'Đã ký nhận (tab NV)']];
      for (const r of f.recs) for (const s of r.ships) day.push([r.date, r.office, s.n, s.sc, s.ok, pct(ratio(s.ok, s.sc)), s.tt, s.tk, pct(ratio(s.tk, s.tt)), s.sg ?? '']);
      addSheet('Shipper theo ngay', day);
    }
  }

  /* ------------------------------------------------------------------ *
   * Tab ANM: so sánh nhiều bưu cục (đọc từ lịch sử đã lưu)
   * ------------------------------------------------------------------ */
  const pct2 = v => v.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const API_SITE_TREE = '/v9-site/site/getAllChildSites';
  // Trang COD dùng mã bưu cục riêng, khác mã của API tìm bưu cục. Thứ tự: mã đã học > mã đã xác nhận > mã của API tìm bưu cục (đoán)
  const COD_IDS_BUILTIN = {};
  function codIds() {
    const m = {}; for (const [k, v] of Object.entries(COD_IDS_BUILTIN)) m[plain(k)] = v;
    try { const o = JSON.parse(localStorage.getItem('b8s_codsites') || '{}'); if (o && typeof o === 'object') for (const [k, v] of Object.entries(o)) if (Number.isFinite(+v)) m[k] = +v; } catch { /* bỏ qua */ }
    return m;
  }
  function saveCodIds(learned) {
    try { const o = JSON.parse(localStorage.getItem('b8s_codsites') || '{}') || {}; localStorage.setItem('b8s_codsites', JSON.stringify(Object.assign(o, learned))); } catch { /* bỏ qua */ }
  }
  const isFS = o => /^\s*fs(\s|$)/i.test(String(o ?? ''));   // bưu cục FS không theo dõi COD
  const isDS = o => /^\s*ds(\s|$)/i.test(String(o ?? ''));   // bản chia sẻ: COD chỉ lấy đúng bưu cục có tên bắt đầu bằng "DS"
  const API_PU_DET = '/v9-order/dc/kanban/am/get-srec-detail';
  const PU_KEEP = 75, PU_TODAY_MS = 8 * 60000, PU_RETRY_MS = 300000, PU_WORKERS = 3;
  const PU_TH = { good: 95, warn: 85 };
  const pkNum = v => { const x = Number(v); return Number.isFinite(x) ? x : null; };
  // Phân loại một dòng đơn
  function pkCls(r) {
    const ty = String(r.customerType ?? '').trim(), pk = !!(r.inNetworkTime), ot = Number(r.isOnTimeRec) === 1;
    const bd = /^bd\s*sale$/i.test(ty), tt = /^tiktok$/i.test(ty);
    const cb = r.isCrossBorder === undefined || r.isCrossBorder === null || r.isCrossBorder === '' ? null : Number(r.isCrossBorder) === 1;
    return { pk, ot, bd, tt, ex: cb !== null ? cb : bd && pk && !ot };   // Kanban loại đơn cờ isCrossBorder = 1 khỏi mẫu số đúng hạn; thiếu trường thì dùng quy tắc cũ
  }
  // Gộp danh sách đơn thành số liệu theo shop (không có TikTok) + số tổng kể cả TikTok để đối chiếu Kanban
  // sh: [tên, loại (B = BD Sale, S = còn lại), mới, cũ, mới đã lấy, cũ đã lấy, mới đúng hạn, mẫu số đúng hạn]
  function pkAggregate(rows) {
    const sh = new Map(), all = { n: 0, o: 0, ot: 0, den: 0, otA: 0, denA: 0 }; let unk = 0;
    for (const r of rows || []) {
      if (!r || typeof r !== 'object') continue;
      const h = Number(r.isHistory);
      if (h !== 0 && h !== 1) { unk++; continue; }
      const c = pkCls(r), nw = h === 0, inDen = !c.ex;
      if (nw) all.n++; else all.o++;
      if (c.ot) all.otA++;
      if (inDen) all.denA++;
      if (nw) { if (c.ot) all.ot++; if (inDen) all.den++; }
      if (c.tt) continue;
      const name = String(r.customerName ?? '').replace(/\s+/g, ' ').trim() || String(r.customerType ?? '').trim() || '(không tên)';
      let s = sh.get(name); if (!s) sh.set(name, s = [name, c.bd ? 'B' : 'S', 0, 0, 0, 0, 0, 0]);
      if (nw) { s[2]++; if (c.pk) s[4]++; if (c.ot) s[6]++; if (inDen) s[7]++; } else { s[3]++; if (c.pk) s[5]++; }
    }
    return { sh: [...sh.values()], all, unk };
  }
  // So với số tổng của Kanban; trả về danh sách điểm lệch (rỗng = khớp)
  function pkVerify(all, srec) {
    const bad = [];
    if (!srec || typeof srec !== 'object') return bad;
    for (const [k, mine, key] of [['đơn mới', all.n, 'todaySrecCnt'], ['đơn cũ', all.o, 'historySrecCnt'], ['đúng hạn đơn mới', all.ot, 'successCnt'], ['mẫu số đúng hạn đơn mới', all.den, 'successDenCnt'], ['đúng hạn tất cả', all.otA, 'timelySrecAhCnt'], ['mẫu số đúng hạn tất cả', all.denA, 'timelySrecAhDenCnt']]) {
      const t = pkNum(srec[key]);
      if (t !== null && t !== mine) bad.push(`${k}: đếm ${mine}, Kanban ${t}`);
    }
    return bad;
  }
  async function pkFetchDay(code, date) {
    const base = { siteCode: code, isSelfData: 0, isCargo: 0, isExpress: 0, nationCode: 'VN', statisticsDate: `${date} 00:00:00` };
    const seen = new Map(); let size = 1000, total = null, got = 0;
    for (let page = 1; page <= 60; page++) {
      let d;
      try { d = await apiPost(API_PU_DET, Object.assign({}, base, { dataKey: 'srecAhCnt', page, size })); }
      catch (e) { if (page === 1 && size > 50 && !isAuthError(e)) { size = size > 200 ? 200 : 50; page--; continue; } throw e; }
      const list = d && Array.isArray(d.list) ? d.list : [];
      if (total === null) total = Number(d && d.total) || 0;
      for (const r of list) { got++; seen.set(r && r.billCode ? String(r.billCode) : `#${got}`, r); }
      if (!list.length || got >= total) break;
    }
    let srec = null;
    try { const s = await apiPost(API_KB_SREC, Object.assign({}, base, { timeFrom: `${date}T00:00:00.000+07:00` })); srec = s && s.srec ? s.srec : null; } catch (e) { if (isAuthError(e)) throw e; }
    return { rows: [...seen.values()], total, srec };
  }
  function pkRec(office, date, got, today) {
    const ag = pkAggregate(got.rows), bad = pkVerify(ag.all, got.srec);
    if (got.total && got.rows.length < got.total) bad.push(`thiếu dòng: nhận ${got.rows.length}/${got.total}`);
    if (ag.unk) bad.push(`${ag.unk} dòng không rõ mới/cũ`);
    return { id: `${office}|${date}`, office, date, at: Date.now(), fin: date < today, sh: ag.sh, bad, v: got.srec ? 1 : 0 };
  }
  // Số tổng của một bản ghi (đã bỏ TikTok)
  function pkSum(rec) {
    const s = { neu: 0, old: 0, pkn: 0, pko: 0, ot: 0, den: 0 };
    for (const x of (rec && rec.sh) || []) { s.neu += x[2]; s.old += x[3]; s.pkn += x[4]; s.pko += x[5]; s.ot += x[6]; s.den += x[7]; }
    return s;
  }
  const pkAdd = (a, b) => { for (const k of Object.keys(a)) a[k] += b[k]; return a; };
  // Tổng của nhiều bưu cục trong một ngày; null nếu thiếu số liệu của bưu cục nào
  function pkTotal(map, offices, date) {
    const t = { neu: 0, old: 0, pkn: 0, pko: 0, ot: 0, den: 0 }, miss = [];
    for (const o of offices) { const r = map.get(`${o}|${date}`); if (!r) miss.push(o); else pkAdd(t, pkSum(r)); }
    return { t: miss.length ? null : t, miss };
  }
  // Gộp shop theo tên qua nhiều bưu cục
  function pkShops(map, offices, date) {
    const m = new Map();
    for (const o of offices) {
      const r = map.get(`${o}|${date}`); if (!r) continue;
      for (const x of r.sh) {
        let s = m.get(x[0]); if (!s) m.set(x[0], s = { name: x[0], t: x[1], neu: 0, old: 0, pkn: 0, pko: 0, ot: 0, den: 0 });
        s.neu += x[2]; s.old += x[3]; s.pkn += x[4]; s.pko += x[5]; s.ot += x[6]; s.den += x[7]; if (x[1] === 'B') s.t = 'B';
      }
    }
    return m;
  }
  // Kỳ thống kê Top 20 shop: từ mùng 1 đến mùng 7 thống kê CẢ THÁNG TRƯỚC, từ mùng 8 thống kê tháng này (mùng 1 đến ngày đang xem)
  function pkTopSpan(date) {
    const s = pkSpan(date);
    if (s.dom <= 7) return { days: pkDays(s.prev0, s.pdim), month: s.pm, prevMonth: true, from: s.prev0, to: addDays(s.prev0, s.pdim - 1) };
    return { days: pkDays(s.cur0, s.dom), month: s.m, prevMonth: false, from: s.cur0, to: date };
  }
  // Top n shop có tổng đơn mới (sản lượng gửi hàng) lớn nhất trong kỳ, gộp theo tên shop qua các bưu cục và các ngày đã có số liệu
  function pkTopMonth(map, offices, days, n = 20) {
    const m = new Map(), dayHave = new Set(); let have = 0, all = 0;
    for (const o of offices) for (const d of days) {
      const r = map.get(`${o}|${d}`); if (!r) continue;
      have++; dayHave.add(d);
      for (const x of r.sh) {
        let t = m.get(x[0]); if (!t) m.set(x[0], t = { name: x[0], t: x[1], neu: 0, pkn: 0, ot: 0, den: 0 });
        t.neu += x[2]; t.pkn += x[4]; t.ot += x[6]; t.den += x[7]; if (x[1] === 'B') t.t = 'B'; all += x[2];
      }
    }
    const list = [...m.values()].filter(t => t.neu > 0).sort((a, b) => b.neu - a.neu || a.name.localeCompare(b.name, 'vi')).slice(0, n);
    return { list, sum: list.reduce((t, x) => t + x.neu, 0), all, have, need: offices.length * days.length, dh: dayHave.size };
  }
  // Khoảng ngày của tháng này (từ mùng 1 đến ngày xem) và tháng trước (cùng số ngày + cả tháng)
  function pkSpan(date) {
    const y = +date.slice(0, 4), m = +date.slice(5, 7), dom = +date.slice(8, 10), py = m === 1 ? y - 1 : y, pm = m === 1 ? 12 : m - 1, pdim = new Date(Date.UTC(py, pm, 0)).getUTCDate();
    return { cur0: `${y}-${pad(m)}-01`, prev0: `${py}-${pad(pm)}-01`, dom, pdim, cut: Math.min(dom, pdim), pm, m };
  }
  const pkDays = (a, n) => Array.from({ length: n }, (_, i) => addDays(a, i));
  function pkMonthDays(date) { const s = pkSpan(date); return pkDays(s.cur0, s.dom).concat(pkDays(s.prev0, s.pdim)); }
  // Tích lũy tháng của một shop ở một bưu cục
  function pkMonth(map, office, name, date) {
    const s = pkSpan(date), cur = pkDays(s.cur0, s.dom), prev = pkDays(s.prev0, s.pdim);
    const pick = d => { const r = map.get(`${office}|${d}`); if (!r) return null; const x = r.sh.find(z => z[0] === name); return x ? { neu: x[2], pkn: x[4], ot: x[6], den: x[7] } : { neu: 0, pkn: 0, ot: 0, den: 0 }; };
    const agg = (days) => { const t = { neu: 0, pkn: 0, ot: 0, den: 0 }; let have = 0; for (const d of days) { const p = pick(d); if (p) { have++; pkAdd(t, p); } } t.have = have; t.need = days.length; return t; };
    return { cur: agg(cur), prevSame: agg(prev.slice(0, s.cut)), prevAll: agg(prev), curSeries: cur.map(d => { const p = pick(d); return p ? p.neu : null; }), prevSeries: prev.map(d => { const p = pick(d); return p ? p.neu : null; }), span: s };
  }
  const pkPct = (a, b) => b > 0 ? `${(a / b * 100).toFixed(2).replace('.', ',')}%` : '—';
  const pkF = n => Math.round(n).toLocaleString('vi-VN');
  const pkDow = k => ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'][new Date(keyToUTC(k)).getUTCDay()];
  const pkDay = k => `${pkDow(k)} ${fmtDM(k)}`;
  function pkPill(a, b) {
    if (!(b > 0)) return '<span class="pk-pill na">không tính</span>';
    const p = a / b * 100;
    return `<span class="pk-pill ${p >= PU_TH.good ? 'good' : p >= PU_TH.warn ? 'warn' : 'bad'}">${p.toFixed(2).replace('.', ',')}%</span>`;
  }
  function pkDelta(v, pts) {
    if (v === null || v === undefined || !Number.isFinite(v)) return '<span class="pk-fl" title="Chưa có số liệu ngày so sánh">—</span>';
    if (pts ? Math.abs(v) < 0.005 : v === 0) return '<span class="pk-fl">0</span>';
    const t = pts ? `${Math.abs(v).toFixed(1).replace('.', ',')} điểm` : pkF(Math.abs(v));
    return `<span class="${v > 0 ? 'pk-up' : 'pk-dn'}">${v > 0 ? '▲ +' : '▼ −'}${t}</span>`;
  }

  // ----- Kho lưu số liệu theo ngày (IndexedDB b8s-pu); ngày đã chốt chỉ tải một lần -----
  function makePkStore(factory) {
    let dbp = null;
    const open = () => dbp || (dbp = new Promise((res, rej) => {
      const rq = factory.open('b8s-pu', 1);
      rq.onupgradeneeded = () => { const db = rq.result; if (!db.objectStoreNames.contains('days')) db.createObjectStore('days', { keyPath: 'id' }); };
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => { dbp = null; rej(rq.error || new Error('Không mở được kho số liệu pickup')); };
      rq.onblocked = () => { dbp = null; rej(new Error('Kho số liệu pickup đang bị tab khác giữ')); };
    }));
    const tx = async (mode, fn) => {
      const db = await open();
      return new Promise((res, rej) => {
        const t = db.transaction('days', mode); let out; const r = fn(t.objectStore('days'));
        if (r) r.onsuccess = () => { out = r.result; };
        t.oncomplete = () => res(out); t.onerror = t.onabort = () => rej(t.error || new Error('Lỗi ghi kho số liệu pickup'));
      });
    };
    return { all: () => tx('readonly', st => st.getAll()), put: r => tx('readwrite', st => { st.put(r); }), delMany: ids => tx('readwrite', st => { ids.forEach(i => st.delete(i)); }) };
  }
  const pkStore = (() => { try { return makePkStore(indexedDB); } catch { return null; } })();
  async function pkLoad() {
    const P = state.pick; if (P.ready) return;
    if (!pkStore) { P.err0 = 'Trình duyệt không hỗ trợ IndexedDB nên tab Pickup-sale không giữ được số liệu các ngày trước (vẫn xem được nhưng mỗi lần mở sẽ tải lại).'; P.ready = true; return; }
    try {
      const lim = addDays(dateKey(new Date()), -PU_KEEP), dead = [];
      for (const r of (await pkStore.all()) || []) {
        if (!r || !r.id || !Array.isArray(r.sh)) { if (r && r.id) dead.push(r.id); continue; }
        if (r.date < lim) dead.push(r.id); else P.map.set(r.id, r);
      }
      if (dead.length) await pkStore.delMany(dead);
    } catch (e) { P.err0 = `Không đọc được kho số liệu pickup: ${e.message}`; }
    P.ready = true;
  }
  const pkOffices = () => anmOffices().filter(o => anmSel().has(o));
  const pkNeed = (id, date, today, force) => {
    const r = state.pick.map.get(id), P = state.pick;
    if (force) return true;
    if (P.tried[id] && Date.now() - P.tried[id] < PU_RETRY_MS) return false;
    if (!r) return true;
    if (r.fin) return false;
    return date < today || Date.now() - r.at > PU_TODAY_MS;
  };
  let pkRedrawT = 0;
  function pkRedraw(now) {
    if (state.view !== 'pick') return;
    if (now) { clearTimeout(pkRedrawT); pkRedrawT = 0; renderPick(); return; }
    if (!pkRedrawT) pkRedrawT = setTimeout(() => { pkRedrawT = 0; if (state.view === 'pick') renderPick(); }, 350);
  }
  // Chạy các việc tải (office, date) với tối đa PU_WORKERS luồng; dừng nếu hết phiên đăng nhập
  async function pkRun(jobs, onProg) {
    const P = state.pick, today = dateKey(new Date()); let i = 0, done = 0, stop = false;
    const worker = async () => {
      while (!stop && i < jobs.length) {
        const j = jobs[i++], id = `${j.office}|${j.date}`;
        P.tried[id] = Date.now();
        try {
          const site = await resolveSite(j.office), got = await pkFetchDay(site.code, j.date), rec = pkRec(j.office, j.date, got, today);
          P.map.set(id, rec); delete P.err[id];
          if (pkStore) { try { await pkStore.put(rec); } catch (e) { P.err0 = `Không lưu được số liệu pickup: ${e.message}`; } }
        } catch (e) {
          P.err[id] = String(e && e.message || e);
          if (isAuthError(e)) { stop = true; P.msg = 'Cần đăng nhập lại trang 800BEST để lấy số liệu pickup.'; }
        }
        done++; if (onProg) onProg(done, jobs.length);
      }
    };
    await Promise.all(Array.from({ length: Math.min(PU_WORKERS, jobs.length) }, worker));
  }
  async function pkEnsure(force) {
    const P = state.pick; await pkLoad();
    if (P.run || !state.fast) { if (!state.fast) { P.msg = 'Đang tắt “Tải nhanh bằng API” trong Cài đặt nên tab Pickup-sale không lấy được số liệu.'; renderPick(); } return; }
    const today = dateKey(new Date()), d = P.date || today, offs = pkOffices(), dates = [d];
    if (P.cmp !== 'w1') dates.push(addDays(d, -1));
    if (P.cmp !== 'd1') dates.push(addDays(d, -7));
    const jobs = []; for (const dt of dates) for (const o of offs) if (pkNeed(`${o}|${dt}`, dt, today, force && dt === d)) jobs.push({ office: o, date: dt });
    if (!jobs.length) { renderPick(); return; }
    P.run = true; P.msg = ''; P.prog = `Đang tải pickup 0/${jobs.length}…`; renderPick();
    try { await pkRun(jobs, (a, b) => { P.prog = `Đang tải pickup ${a}/${b}…`; pkRedraw(); }); }
    finally { P.run = false; P.prog = ''; P.at = Date.now(); renderPick(); }
  }
  // Tải các ngày còn thiếu của kỳ Top 20 (cả tháng) cho các bưu cục đang chọn; ngày đã chốt chỉ tải một lần và giữ trong kho
  async function pkEnsureTop() {
    const P = state.pick; await pkLoad();
    if (P.run || !state.fast) { if (!state.fast) { P.msg = 'Đang tắt “Tải nhanh bằng API” trong Cài đặt nên tab Pickup-sale không lấy được số liệu.'; renderPick(); } return; }
    const today = dateKey(new Date()), D = P.date || today, sp = pkTopSpan(D), offs = pkOffices();
    const jobs = []; for (const o of offs) for (const dt of sp.days) if (dt <= today && pkNeed(`${o}|${dt}`, dt, today, false)) jobs.push({ office: o, date: dt });
    if (!jobs.length) { renderPick(); return; }
    P.run = true; P.msg = ''; P.prog = `Đang tải số liệu tháng ${sp.month} cho Top 20: 0/${jobs.length}…`; renderPick();
    try { await pkRun(jobs, (a, b) => { P.prog = `Đang tải số liệu tháng ${sp.month} cho Top 20: ${a}/${b}…`; pkRedraw(); }); }
    finally { P.run = false; P.prog = ''; P.at = Date.now(); renderPick(); }
  }
  async function pkEnsureMonth(office) {
    const P = state.pick, today = dateKey(new Date()), d = P.date || today;
    if (!office || P.mon.run || !state.fast) return;
    const jobs = pkMonthDays(d).filter(dt => dt <= today && pkNeed(`${office}|${dt}`, dt, today, false)).map(dt => ({ office, date: dt }));
    if (!jobs.length) return;
    P.mon = { run: true, done: 0, total: jobs.length, office };
    try { await pkRun(jobs, (a, b) => { P.mon.done = a; pkRedraw(); }); }
    finally { P.mon = { run: false, done: 0, total: 0, office: '' }; pkRedraw(true); }
  }

  // ----- Vẽ tab -----
  function pkChart(cur, prev, span) {
    const W = 760, H = 190, L = 36, R = 40, T = 10, B = 22, n = Math.max(span.pdim, span.dom, 28);
    let max = 0; for (const v of cur.concat(prev)) if (v !== null && v > max) max = v;
    max = Math.max(10, Math.ceil(max * 1.1 / 10) * 10);
    const X = d => L + (W - L - R) * (d - 1) / (n - 1), Y = v => T + (H - T - B) * (1 - v / max);
    let g = '';
    for (const q of [0, .5, 1]) { const v = Math.round(max * q); g += `<line x1="${L}" x2="${W - R}" y1="${Y(v)}" y2="${Y(v)}" stroke="#e5e7eb"/><text x="${L - 6}" y="${Y(v) + 4}" text-anchor="end" font-size="10" fill="#6b7280">${v}</text>`; }
    for (const d of [1, 5, 10, 15, 20, 25, 30].filter(x => x <= n)) g += `<text x="${X(d)}" y="${H - 6}" text-anchor="middle" font-size="10" fill="#6b7280">${d}</text>`;
    const line = (a, c, dash) => { let p = '', pen = false; a.forEach((v, i) => { if (v === null) { pen = false; return; } p += `${pen ? 'L' : 'M'}${X(i + 1).toFixed(1)} ${Y(v).toFixed(1)}`; pen = true; }); return `<path d="${p}" fill="none" stroke="${c}" stroke-width="2" ${dash ? 'stroke-dasharray="5 4"' : ''} stroke-linejoin="round"/>`; };
    let last = -1; cur.forEach((v, i) => { if (v !== null) last = i; });
    const end = last >= 0 ? `<circle cx="${X(last + 1)}" cy="${Y(cur[last])}" r="4" fill="#4c3a8f"/><text x="${X(last + 1) + 7}" y="${Y(cur[last]) + 4}" font-size="11" font-weight="700" fill="#4c3a8f">${cur[last]}</text>` : '';
    return `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="Đơn mới mỗi ngày, tháng này so với tháng trước">${g}${line(prev, '#d97706', true)}${line(cur, '#4c3a8f', false)}${end}</svg>`;
  }
  const pkTag = t => t === 'B' ? '<span class="pk-tag bd">BD Sale</span>' : '<span class="pk-tag">Shop</span>';
  function renderPick() {
    const box = $('#b8s-pick'); if (!box) return;
    const P = state.pick, now = new Date(), today = dateKey(now);
    if (!anmOffices().length) { box.innerHTML = '<div class="pk-box" style="text-align:center;padding:28px"><b>Chưa có bưu cục nào để theo dõi</b><div class="pk-mu">Bấm “Tải số liệu” ở tab ANM để tool tự tìm bưu cục, hoặc thêm ở Cài đặt.</div></div>'; return; }
    const D = P.date || today, D1 = addDays(D, -1), W1 = addDays(D, -7), offs = pkOffices(), allOffs = anmOffices();
    const cmpD = []; if (P.cmp !== 'w1') cmpD.push(['D-1', D1]); if (P.cmp !== 'd1') cmpD.push(['W-1', W1]);
    const cur = pkTotal(state.pick.map, offs, D), have = offs.filter(o => P.map.has(`${o}|${D}`));
    if (P.off && !offs.includes(P.off)) { P.off = ''; P.shop = ''; }
    if (!P.off) P.off = have[0] || offs[0] || '';
    const T0 = cur.t || have.reduce((t, o) => pkAdd(t, pkSum(P.map.get(`${o}|${D}`))), { neu: 0, old: 0, pkn: 0, pko: 0, ot: 0, den: 0 });
    const dd = (fn, key) => cmpD.map(([lb, dt]) => { const c = pkTotal(P.map, offs, dt); return `<em>${lb}</em> ${c.t && cur.t ? pkDelta(fn(T0) - fn(c.t), key === 'pts') : '<span class="pk-fl" title="Chưa đủ số liệu ngày so sánh">—</span>'}`; }).join(' ');
    const sp = pkTopSpan(D), top = pkTopMonth(P.map, offs, sp.days), topFull = top.have >= top.need;
    const tot = T0.neu + T0.old, got = T0.pkn + T0.pko;
    const card = (lb, v, sub, d, extra) => `<div class="pk-c${extra ? ' click' : ''}"${extra || ''}><span>${lb}</span><b>${v}</b><small>${sub}</small><div class="pk-d">${d}</div></div>`;
    const cards = [
      card('Tổng đơn cần PU', pkF(tot), `${pkF(T0.neu)} mới + ${pkF(T0.old)} cũ`, dd(t => t.neu + t.old)),
      card('Đơn mới', pkF(T0.neu), `${offs.length} bưu cục`, dd(t => t.neu)),
      card('Đơn cũ (tồn)', pkF(T0.old), `${pkF(T0.old - T0.pko)} chưa PU`, dd(t => t.old)),
      card('PU thành công', pkF(got), `${pkF(T0.pkn)} mới (${pkPct(T0.pkn, T0.neu)}) · ${pkF(T0.pko)} cũ`, dd(t => t.pkn + t.pko)),
      card('Chưa PU', pkF(tot - got), `${pkF(T0.neu - T0.pkn)} mới · ${pkF(T0.old - T0.pko)} cũ`, dd(t => (t.neu + t.old) - (t.pkn + t.pko))),
      card('PU đúng hạn, đơn mới', pkPct(T0.ot, T0.den), `${pkF(T0.ot)}/${pkF(T0.den)}`, dd(t => t.den ? t.ot / t.den * 100 : 0, 'pts')),
      card(`Top 20 shop · tháng ${sp.month}`, top.have ? pkF(top.sum) : '—', topFull ? `chiếm ${pkPct(top.sum, top.all)} đơn mới cả tháng · <b class="go">bấm xem</b>` : `mới có ${top.have}/${top.need} lượt ngày · <b class="go">bấm để tải đủ</b>`, `<span class="pk-mu">${fmtDM(sp.from)} → ${fmtDM(sp.to)}</span>`, ' data-a="pkTop" title="Bấm để xem 20 shop có sản lượng lớn nhất trong tháng"')
    ].join('');
    // cảnh báo
    const al = []; const oldU = [], exs = [];
    for (const o of have) for (const x of P.map.get(`${o}|${D}`).sh) { if (x[3] - x[5] > 0) oldU.push({ o, n: x[0], v: x[3] - x[5] }); if (x[2] - x[7] > 0) exs.push({ o, n: x[0], v: x[2] - x[7] }); }
    const oldTot = oldU.reduce((t, x) => t + x.v, 0);
    if (oldTot > 0) { const m = oldU.reduce((a, b) => b.v > a.v ? b : a), ob = new Map(); oldU.forEach(x => ob.set(x.o, (ob.get(x.o) || 0) + x.v)); const [bo, bv] = [...ob.entries()].sort((a, b) => b[1] - a[1])[0];
      al.push(`<div class="pk-al ${oldTot >= 100 ? 'bad' : 'warn'}"><b>Tồn cũ:</b> ${pkF(oldTot)} đơn cũ chưa PU toàn khu vực, nhiều nhất ở <b>${esc(bo)}</b> (${pkF(bv)} đơn). Shop nhiều nhất: <b>${esc(m.n)}</b> (${esc(m.o)}, ${pkF(m.v)} đơn).</div>`); }
    if (exs.length) { const m = exs.reduce((a, b) => b.v > a.v ? b : a); al.push(`<div class="pk-al warn"><b>${esc(m.n)}</b> tạo ${pkF(m.v)} đơn tại điểm lấy (${esc(m.o)}), không tính vào đúng hạn.</div>`); }
    const lows = have.map(o => { const s = pkSum(P.map.get(`${o}|${D}`)); return { o, r: s.neu ? s.pkn / s.neu : null }; }).filter(x => x.r !== null).sort((a, b) => a.r - b.r);
    if (lows.length > 1) al.push(`<div class="pk-al info"><b>PU đơn mới thấp nhất:</b> ${esc(lows[0].o)} ${pkPct(lows[0].r, 1)}</div>`);
    // bảng bưu cục
    const orow = o => {
      const r = P.map.get(`${o}|${D}`), er = P.err[`${o}|${D}`];
      if (!r) return `<tr data-a="pkOff" data-o="${esc(o)}" class="${o === P.off ? 'sel' : ''}"><td class="nm"><b>${esc(o)}</b></td><td colspan="7" class="pk-mu">${er ? '⚠ không tải được: ' + esc(er) : P.run ? 'đang tải…' : 'chưa tải'}</td></tr>`;
      const s = pkSum(r), tt = s.neu + s.old, g = s.pkn + s.pko, w1 = tt ? s.pkn / tt * 100 : 0, w2 = tt ? s.pko / tt * 100 : 0;
      return `<tr data-a="pkOff" data-o="${esc(o)}" class="${o === P.off ? 'sel' : ''}"><td class="nm"><b>${esc(o)}</b>${r.bad.length ? ` <span class="pk-warn" title="${esc('Lệch với Kanban: ' + r.bad.join('; '))}">⚠</span>` : ''}</td><td class="n">${pkF(tt)}</td>
        <td><div class="pk-stack" title="${pkF(s.pkn)} PU mới · ${pkF(s.pko)} PU cũ · ${pkF(tt - g)} chưa PU"><i style="width:${w1.toFixed(2)}%;background:#4c3a8f"></i><i style="width:${w2.toFixed(2)}%;background:#d97706"></i></div></td>
        <td class="n">${pkF(s.neu)}</td><td class="n">${pkF(s.old)}</td><td class="n"><b>${pkF(s.pkn)}/${pkF(s.neu)}</b> ${pkPill(s.pkn, s.neu)}</td><td class="n"><b>${pkF(s.pko)}/${pkF(s.old)}</b> <span class="pk-mu">${pkPct(s.pko, s.old)}</span></td><td class="n"><b>${pkF(g)}/${pkF(tt)}</b> <span class="pk-mu">${pkPct(g, tt)}</span></td></tr>`;
    };
    // bảng shop của bưu cục đang chọn
    const rec = P.map.get(`${P.off}|${D}`), recC = cmpD.map(([lb, dt]) => ({ lb, dt, r: P.map.get(`${P.off}|${dt}`) }));
    const shopRows = rec ? rec.sh.slice().sort((a, b) => (b[2] + b[3]) - (a[2] + a[3]) || a[0].localeCompare(b[0], 'vi')) : [];
    const cmpCell = (c, name, neu) => { if (!c.r) return '<td class="n"><span class="pk-fl" title="Chưa có số liệu ngày so sánh">—</span></td>'; const x = c.r.sh.find(z => z[0] === name); return `<td class="n">${pkDelta(neu - (x ? x[2] : 0))}</td>`; };
    const srow = x => {
      const tt = x[2] + x[3], g = x[4] + x[5];
      return `<tr data-a="pkShop" data-s="${esc(x[0])}" class="${x[0] === P.shop ? 'sel' : ''}"><td class="nm" title="${esc(x[0])}">${esc(x[0])}${pkTag(x[1])}</td><td class="n">${pkF(tt)}</td><td class="n">${pkF(x[3])}</td><td class="n"><b>${pkF(x[2])}</b></td>${recC.map(c => cmpCell(c, x[0], x[2])).join('')}<td class="n">${pkF(x[4])}/${pkF(x[2])}</td><td class="n">${pkF(x[5])}/${pkF(x[3])}</td><td class="n"><b>${pkF(g)}/${pkF(tt)}</b> <span class="pk-mu">${pkPct(g, tt)}</span></td><td class="n">${pkPill(x[6], x[7])}</td></tr>`;
    };
    // chi tiết shop (chỉ hiện khi bấm một shop)
    let detail = '';
    const sx = P.shop && rec ? rec.sh.find(z => z[0] === P.shop) : null;
    if (sx) {
      const m = pkMonth(P.map, P.off, P.shop, D), c = m.cur, p = m.prevSame, a = m.prevAll, dlt = c.neu - p.neu, dp = p.neu ? dlt / p.neu * 100 : 0;
      const partial = c.have < c.need || p.have < p.need || a.have < a.need;
      const mon = P.mon.run && P.mon.office === P.off ? `<span class="pk-mu"> · đang tải ngày còn thiếu ${P.mon.done}/${P.mon.total}…</span>` : (partial ? `<span class="pk-warn"> · mới có ${c.have}/${c.need} ngày tháng này, ${a.have}/${a.need} ngày tháng trước</span>` : '');
      const mt = (k, v, s, cls) => `<div class="pk-m${cls ? ' ' + cls : ''}"><span class="k">${k}</span><span class="v">${v}</span><span class="s">${s}</span></div>`;
      detail = `<section class="pk-box" id="pk-detail"><header><h2>Thống kê shop · ${esc(P.shop)} ${pkTag(sx[1])}<small class="pk-mu"> · ${esc(P.off)}</small>${mon}</h2><a href="#" data-a="pkShopClose" class="pk-x">Đóng ×</a></header>
        <div class="pk-low"><div class="pk-tiles">
          ${mt(`Tháng ${m.span.m} (1 → ${fmtDM(D)})`, `${pkF(c.neu)} đơn mới`, `PU ${pkF(c.pkn)} (${pkPct(c.pkn, c.neu)}) · đúng hạn ${pkPct(c.ot, c.den)}`, 'cur')}
          ${mt(`Tháng ${m.span.pm} cùng kỳ (1 → ${pad(m.span.cut)}/${pad(m.span.pm)})`, `${pkF(p.neu)} đơn mới`, `PU ${pkF(p.pkn)} (${pkPct(p.pkn, p.neu)}) · đúng hạn ${pkPct(p.ot, p.den)}`)}
          ${mt('Chênh lệch cùng kỳ', `<span class="${dlt > 0 ? 'pk-up' : dlt < 0 ? 'pk-dn' : 'pk-fl'}">${dlt > 0 ? '▲ +' : dlt < 0 ? '▼ −' : ''}${pkF(Math.abs(dlt))} đơn</span>`, p.neu ? `${dlt > 0 ? '+' : dlt < 0 ? '−' : ''}${Math.abs(dp).toFixed(1).replace('.', ',')}% so với tháng ${m.span.pm}` : `tháng ${m.span.pm} chưa có đơn`)}
          ${mt(`Tổng đơn cả tháng ${m.span.pm}`, `${pkF(a.neu)} đơn`, `tháng ${m.span.m} mới đạt ${pkPct(c.neu, a.neu)} của con số này`)}
        </div><div><div class="pk-leg"><span style="--k:#4c3a8f">Tháng ${m.span.m}</span><span style="--k:#d97706">Tháng ${m.span.pm}</span><span class="pk-mu">đơn mới mỗi ngày</span></div>${pkChart(m.curSeries, m.prevSeries, m.span)}</div></div></section>`;
    }
    // Top 20 shop theo sản lượng (tổng đơn mới) cả tháng: mùng 1 đến mùng 7 thống kê tháng trước, từ mùng 8 thống kê tháng này
    let top20 = '';
    if (P.top) {
      const tl = top.list, S = tl.reduce((t, x) => { t.neu += x.neu; t.pkn += x.pkn; t.ot += x.ot; t.den += x.den; return t; }, { neu: 0, pkn: 0, ot: 0, den: 0 });
      const tile = (lb, v, sb, cls) => `<div class="pk-m${cls ? ' ' + cls : ''}"><span class="k">${lb}</span><span class="v">${v}</span><span class="s">${sb}</span></div>`;
      const rowT = (x, i) => `<tr><td class="pk-mu">${i + 1}</td><td class="nm" title="${esc(x.name)}">${esc(x.name)}${pkTag(x.t)}</td><td class="n"><b>${pkF(x.neu)}</b></td><td class="n">${pkF(x.neu / Math.max(1, top.dh))}</td><td class="n">${pkF(x.pkn)}/${pkF(x.neu)} ${pkPill(x.pkn, x.neu)}</td><td class="n">${pkPct(x.ot, x.den)}</td><td class="n">${pkPct(x.neu, top.all)}</td></tr>`;
      top20 = `<section class="pk-box" id="pk-top"><header><h2>3. Top 20 shop sản lượng lớn nhất · tháng ${sp.month} (${fmtDM(sp.from)} → ${fmtDM(sp.to)}) · ${offs.length} bưu cục</h2><a href="#" data-a="pkTopClose" class="pk-x">Đóng ×</a></header>
        <div class="pk-al info">${sp.prevMonth ? `Đang là đầu tháng (mùng 1 đến mùng 7) nên thống kê cả tháng ${sp.month}. Từ mùng 8 sẽ chuyển sang tháng mới.` : `Thống kê từ mùng 1 đến ${fmtDM(sp.to)}. Từ mùng 1 đến mùng 7 tháng sau sẽ chuyển sang thống kê cả tháng ${sp.month}.`} Xếp theo tổng đơn mới, không tính TikTok.</div>
        ${topFull ? '' : `<div class="pk-al warn">Mới có ${top.have}/${top.need} lượt ngày × bưu cục nên top 20 chưa đủ. ${P.run ? 'Đang tải…' : '<a href="#" data-a="pkTopLoad">Bấm để tải phần còn thiếu</a>'}</div>`}
        <div class="pk-tiles4">${tile('Tổng đơn mới 20 shop', pkF(S.neu), `chiếm ${pkPct(S.neu, top.all)} đơn mới của ${offs.length} bưu cục`, 'cur')}${tile('Đơn mới cả khu vực', pkF(top.all), `${top.dh} ngày có số liệu`)}${tile('PU đơn mới của 20 shop', pkPct(S.pkn, S.neu), `${pkF(S.pkn)}/${pkF(S.neu)}`)}${tile('PU đúng hạn, đơn mới', pkPct(S.ot, S.den), `${pkF(S.ot)}/${pkF(S.den)}`)}</div>
        <div class="pk-sc"><table><thead><tr><th>#</th><th>Shop</th><th class="n">Đơn mới trong tháng</th><th class="n">TB/ngày</th><th class="n">PU đơn mới</th><th class="n">PU đúng hạn</th><th class="n">Tỉ trọng đơn mới</th></tr></thead><tbody>${tl.length ? tl.map(rowT).join('') : '<tr><td colspan="7" class="pk-mu">Chưa có số liệu.</td></tr>'}</tbody></table></div></section>`;
    }
    const bads = have.filter(o => P.map.get(`${o}|${D}`).bad.length).map(o => `${o}: ${P.map.get(`${o}|${D}`).bad.join('; ')}`);
    const dt = new Date(keyToUTC(D) + 12 * 3600000), isToday = D >= today;
    const status = `Đang xem: <b>${pkDay(D)}/${D.slice(0, 4)}</b>${cmpD.length ? ' · so với ' + cmpD.map(([lb, d]) => `<b>${lb} (${pkDay(d)})</b>`).join(' và ') : ''} · không tính đơn TikTok · ${have.length}/${offs.length} bưu cục có số liệu${P.prog ? ' · ' + esc(P.prog) : ''}${!P.run && P.at ? ` · cập nhật ${pad(new Date(P.at).getHours())}:${pad(new Date(P.at).getMinutes())}` : ''}`;
    const errs = offs.filter(o => P.err[`${o}|${D}`] && !P.map.has(`${o}|${D}`));
    box.innerHTML = `<div class="pk">
      <div class="pk-ctl"><span class="pk-chip">Bưu cục: ${offs.length}/${allOffs.length} đã chọn <small>(đổi ở tab ANM)</small></span><label>Ngày <input type="date" id="pk-date" value="${esc(D)}" max="${esc(today)}"></label>
        <label>So sánh <select id="pk-cmp"><option value="both" ${P.cmp === 'both' ? 'selected' : ''}>D-1 và W-1</option><option value="d1" ${P.cmp === 'd1' ? 'selected' : ''}>Chỉ D-1</option><option value="w1" ${P.cmp === 'w1' ? 'selected' : ''}>Chỉ W-1</option></select></label>
        <span style="flex:1"></span><button class="pk-go" data-a="pkLoad" ${P.run ? 'disabled' : ''}>${P.run ? '⟳ đang tải…' : `⟳ Tải số liệu ${offs.length} bưu cục`}</button></div>
      <div class="pk-st">${status}</div>
      ${P.msg ? `<div class="pk-al bad">🔒 ${esc(P.msg)}</div>` : ''}${P.err0 ? `<div class="pk-al warn">${esc(P.err0)}</div>` : ''}${errs.length ? `<div class="pk-al bad">Không tải được: ${esc(errs.slice(0, 3).map(o => `${o} (${P.err[`${o}|${D}`]})`).join('; '))}${errs.length > 3 ? '…' : ''}</div>` : ''}${bads.length ? `<div class="pk-al warn">⚠ Số đếm lệch số tổng của Kanban, nên đối chiếu: ${esc(bads.join(' | '))}</div>` : ''}
      <div class="pk-cards">${cards}</div>
      ${al.length ? `<div class="pk-als">${al.join('')}</div>` : ''}
      <section class="pk-box"><header><h2>1. Chỉ số lấy hàng theo bưu cục</h2><span class="pk-mu">Bấm tên bưu cục để xem shop bên dưới</span></header>
        <div class="pk-sc"><table><thead><tr><th>Bưu cục</th><th class="n">Tổng</th><th>Cơ cấu <span class="pk-mu">(PU mới · PU cũ · chưa PU)</span></th><th class="n">Mới</th><th class="n">Cũ</th><th class="n">PU đơn mới</th><th class="n">PU đơn cũ</th><th class="n">Tổng PU thành công / tổng cần PU</th></tr></thead><tbody>${offs.map(orow).join('')}</tbody></table></div></section>
      <section class="pk-box"><header><h2>2. Theo shop · ${esc(P.off)}</h2><span class="pk-mu">Bấm một shop để xem thống kê của shop đó bên dưới</span></header>
        <div class="pk-sc"><table><thead><tr><th>Shop</th><th class="n">Cần PU</th><th class="n">Cũ</th><th class="n">Mới</th>${cmpD.map(([lb]) => `<th class="n">Mới so ${lb}</th>`).join('')}<th class="n">PU mới</th><th class="n">PU cũ</th><th class="n">PU thành công / tổng</th><th class="n">Đúng hạn</th></tr></thead><tbody>${shopRows.map(srow).join('') || `<tr><td colspan="${8 + cmpD.length}" class="pk-mu">${P.run ? 'đang tải…' : 'Chưa có số liệu của bưu cục này.'}</td></tr>`}</tbody></table></div></section>
      ${detail}${top20}
      <div class="pk-foot">ⓘ Mới/cũ theo ngày tạo đơn · PU = đã lấy (có giờ vào mạng) · đúng hạn bỏ đơn có cờ isCrossBorder (cách Kanban tính) · tháng = cộng đơn mới từng ngày, mỗi ngày chỉ tải một lần rồi lưu · không tính TikTok.</div></div>`;
  }
  async function onEnterPick() {
    const P = state.pick; renderPick(); await pkEnsure(false);
    if (P.shop) pkEnsureMonth(P.off).catch(() => {});
  }
  function onPickClick(e) {
    const el = e.target.closest('[data-a]'); if (!el) return;
    const P = state.pick, a = el.dataset.a;
    if (a === 'pkOff') { P.off = el.dataset.o; P.shop = ''; renderPick(); }
    else if (a === 'pkShop') { P.shop = P.shop === el.dataset.s ? '' : el.dataset.s; renderPick(); if (P.shop) { pkEnsureMonth(P.off).catch(() => {}); const d = $('#pk-detail'); if (d && d.scrollIntoView) d.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); } }
    else if (a === 'pkShopClose') { e.preventDefault(); P.shop = ''; renderPick(); }
    else if (a === 'pkTop') { P.top = true; renderPick(); pkEnsureTop().catch(er => { P.msg = String(er && er.message || er); renderPick(); }); const d = $('#pk-top'); if (d && d.scrollIntoView) d.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }
    else if (a === 'pkTopLoad') { e.preventDefault(); pkEnsureTop().catch(er => { P.msg = String(er && er.message || er); renderPick(); }); }
    else if (a === 'pkTopClose') { e.preventDefault(); P.top = false; renderPick(); }
    else if (a === 'pkLoad') { pkEnsure(true).then(() => { if (P.shop) pkEnsureMonth(P.off); }).catch(() => {}); }
  }
  function onPickChange(e) {
    const P = state.pick, t = e.target;
    if (t.id === 'pk-date') { if (/^\d{4}-\d{2}-\d{2}$/.test(t.value)) { P.date = t.value >= dateKey(new Date()) ? '' : t.value; P.shop = ''; renderPick(); pkEnsure(false); } }
    else if (t.id === 'pk-cmp') { P.cmp = t.value; renderPick(); pkEnsure(false); }
  }

  function b5Actual(recs) {
    if (!Array.isArray(recs) || !recs.length) return null;
    const dates = recs.map(r => r && r.date).filter(Boolean).sort(), last = dates[dates.length - 1];
    let sum = 0, n = 0;
    for (const r of recs) if (r && r.date === last && r.backlog && Number.isFinite(r.backlog.d5)) { sum += r.backlog.d5; n++; }
    return n ? sum : null;
  }
  const b5Color = v => v > 20 ? '#dc2626' : v >= 10 ? '#d97706' : '#16a34a';
  // rows: d.rows của tab ANM (đã chỉ gồm bưu cục đang tick); trả về danh sách xếp từ nhiều đến ít, bưu cục chưa có số để cuối
  function b5Rows(rows) {
    const list = (rows || []).map(r => ({ o: r.office, v: b5Actual(r.recs) }));
    list.sort((a, b) => (a.v === null) - (b.v === null) || (b.v ?? 0) - (a.v ?? 0) || a.o.localeCompare(b.o, 'vi'));
    const have = list.filter(x => x.v !== null);
    return { list, total: have.length ? have.reduce((t, x) => t + x.v, 0) : null, miss: list.length - have.length };
  }
  function b5Canvas(res, title, sub) {
    const S = 2, W = 560, PADX = 18, HEAD = 70, TH = 36, ROW = 50, FOOT = 30, n = res.list.length + 1;
    const H = HEAD + TH + ROW * n + FOOT, c = document.createElement('canvas');
    c.width = W * S; c.height = H * S;
    const g = c.getContext('2d'); g.scale(S, S);
    const F = '"Segoe UI",Roboto,Arial,sans-serif';
    const rr = (x, y, w, h, r) => { g.beginPath(); g.moveTo(x + r, y); g.arcTo(x + w, y, x + w, y + h, r); g.arcTo(x + w, y + h, x, y + h, r); g.arcTo(x, y + h, x, y, r); g.arcTo(x, y, x + w, y, r); g.closePath(); };
    g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
    g.save(); rr(1, 1, W - 2, H - 2, 14); g.clip();
    const gr = g.createLinearGradient(0, 0, W, HEAD); gr.addColorStop(0, '#667eea'); gr.addColorStop(1, '#764ba2');
    g.fillStyle = gr; g.fillRect(0, 0, W, HEAD);
    g.textBaseline = 'alphabetic'; g.fillStyle = '#ffe58a'; g.font = `700 24px ${F}`; g.fillText(title, PADX, 34);
    g.fillStyle = '#ffffff'; g.font = `400 14px ${F}`; g.fillText(sub, PADX, 56);
    g.fillStyle = '#eef0ff'; g.fillRect(0, HEAD, W, TH);
    g.fillStyle = '#4c3a8f'; g.font = `700 14px ${F}`; g.textAlign = 'left'; g.fillText('Bưu cục', PADX, HEAD + 23);
    g.textAlign = 'right'; g.fillText('Đơn backlog', W - PADX, HEAD + 23);
    res.list.forEach((x, i) => {
      const y = HEAD + TH + ROW * i;
      g.strokeStyle = '#e5e7eb'; g.lineWidth = 1; g.beginPath(); g.moveTo(0, y + ROW - .5); g.lineTo(W, y + ROW - .5); g.stroke();
      g.textAlign = 'left'; g.fillStyle = '#1f2937'; g.font = `600 19px ${F}`; g.fillText(x.o, PADX, y + 32);
      g.textAlign = 'right'; g.font = `800 25px ${F}`;
      g.fillStyle = x.v === null ? '#9ca3af' : b5Color(x.v); g.fillText(x.v === null ? '—' : String(Math.round(x.v)), W - PADX, y + 34);
    });
    const ty = HEAD + TH + ROW * res.list.length;
    g.fillStyle = '#f3f0ff'; g.fillRect(0, ty, W, ROW); g.fillStyle = '#4c3a8f'; g.fillRect(0, ty, W, 2);
    g.textAlign = 'left'; g.fillStyle = '#1f2937'; g.font = `800 19px ${F}`; g.fillText('TỔNG', PADX, ty + 32);
    g.textAlign = 'right'; g.font = `800 25px ${F}`; g.fillText(res.total === null ? '—' : String(Math.round(res.total)), W - PADX, ty + 34);
    g.fillStyle = '#fafafa'; g.fillRect(0, ty + ROW, W, FOOT);
    g.textAlign = 'left'; g.fillStyle = '#6b7280'; g.font = `400 12px ${F}`; g.fillText('Đơn chưa giao thành công, ngày cần phát cách hôm nay từ 5 ngày trở lên', PADX, ty + ROW + 20);
    g.restore();
    g.strokeStyle = '#4c3a8f'; g.lineWidth = 2; rr(1, 1, W - 2, H - 2, 14); g.stroke();
    return c;
  }
  async function b5Shot() {
    const a = state.anm, d = anmData(), now = new Date();
    if (!d || !d.rows || !d.rows.length) { log('Chưa có số liệu để chụp. Bấm “Tải số liệu” trước.', true); return; }
    const res = b5Rows(d.rows);
    if (res.total === null) { log('Chưa có số Backlog 5 ngày của bưu cục nào để chụp.', true); return; }
    const c = b5Canvas(res, 'Backlog 5 ngày', `Khu vực ${REGION} · ${pad(now.getHours())}:${pad(now.getMinutes())} ${fmtDMY(dateKey(now))}`);
    const blob = await new Promise(r => c.toBlob(r, 'image/png'));
    if (!blob) { log('Không tạo được ảnh.', true); return; }
    let ok = false;
    try { await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]); ok = true; } catch { /* tải về */ }
    if (!ok) {
      try { const u = URL.createObjectURL(blob), el = document.createElement('a'); el.href = u; el.download = `Backlog5ngay_${dateKey(now)}.png`; document.body.appendChild(el); el.click(); el.remove(); setTimeout(() => URL.revokeObjectURL(u), 4000); a.shotMsg = 'Trình duyệt không cho copy ảnh, đã tải ảnh về máy.'; }
      catch { a.shotMsg = 'Không copy được ảnh.'; }
    } else a.shotMsg = 'Đã copy ảnh, mở Zalo bấm Ctrl+V để dán.';
    a.shotAt = Date.now(); log(a.shotMsg, !ok && /Không copy/.test(a.shotMsg));
    setTimeout(() => { if (state.view === 'anm') renderAnm(); }, 4200);
    renderAnm();
  }

  /* ------------------------------------------------------------------ *
   * v3.28: danh sách đơn BACKLOG 5 NGÀY từng bưu cục (ngày cần phát cách hôm nay từ 5 ngày trở lên và chưa giao thành công,
   * đúng định nghĩa của cột “Backlog 5 ngày”; cột “Trễ” cho biết số ngày quá hạn).
   * Chỉ tải khi bấm nút. Mỗi bưu cục giữ MỘT danh sách (IndexedDB b8s-bl): tải lại sẽ GHI ĐÈ danh sách cũ.
   * ------------------------------------------------------------------ */
  const BL_MIN = 5;   // chỉ lấy đơn backlog từ 5 ngày trở lên (đúng cột “Backlog 5 ngày”)
  // Địa chỉ giao: đọc từ nhật ký hoạt động của đơn (cột “Ghi chú” của dòng sortingCode) qua billBasicInfo, đúng nội dung trang tra cứu đơn đang hiện cho tài khoản này.
  // Chỉ giữ trong bộ nhớ của phiên (không lưu vào IndexedDB). KHÔNG đọc số điện thoại, tên người nhận, không giải mã/mở che.
  const BL_API_BILL = '/v9-order/tracking/billBasicInfo';
  const BLA = { map: new Map(), run: null, err: '' };
  function blAddrOf(logs) {
    const arr = Array.isArray(logs) ? logs : [];
    for (let i = arr.length - 1; i >= 0; i--) {
      const l = arr[i];
      if (!l || l.operationName !== 'sortingCode') continue;
      const u = String(l.udf4 ?? '').trim();
      if (u) return u.normalize('NFC');
      const m = String(l.reMark ?? '').split(';')[0];
      if (m.trim()) return m.replace(/,+/g, ', ').replace(/\s+/g, ' ').trim().normalize('NFC');
    }
    return '';
  }
  async function blAddrFetch(office) {
    const rec = BL.map.get(office); if (!rec || BLA.run || BL.run) return;
    const codes = rec.items.map(x => x[0]).filter(Boolean);
    BLA.run = { o: office, i: 0, n: codes.length }; BLA.err = ''; blRender();
    let p = 0, bad = 0, auth = false;
    const worker = async () => {
      while (p < codes.length && !auth) {
        const c = codes[p++];
        try {
          const d = await apiPost(BL_API_BILL, { billCode: c }), b = (d && d.taoBaoBillVo) || {};
          BLA.map.set(c, { addr: blAddrOf(d && d.taobaoStatusLogVos), area: [b.consigneeCity, b.consigneeProv].map(x => String(x ?? '').trim()).filter(Boolean).join(', ') });
        } catch (e) { bad++; if (isAuthError(e)) auth = true; }
        BLA.run.i++; if (BLA.run.i % 4 === 0) blRender();
      }
    };
    try { await Promise.all([worker(), worker(), worker(), worker()]); }
    finally { BLA.run = null; BLA.err = auth ? 'Cần đăng nhập lại trang 800BEST để lấy địa chỉ.' : bad ? `Chưa lấy được địa chỉ của ${bad}/${codes.length} đơn.` : ''; blRender(); }
  }
  const blHasAddr = rec => !!rec && rec.items.some(x => BLA.map.has(x[0]));
  const blAddrText = a => a ? ([a.addr, a.area && !(a.addr || '').includes(a.area.split(',')[0]) ? a.area : ''].filter(Boolean).join(' · ') || a.area || '') : '';
  const BL = { map: new Map(), ready: false, run: null, open: '', msg: '', err: '', copied: '', show: false };
  // Hiện trong cửa sổ nổi (không chiếm chỗ trong trang ANM để trang vẫn vừa một màn hình)
  function blRender() {
    let ov = document.getElementById('b8s-bl');
    if (!BL.show) { if (ov) ov.remove(); return; }
    if (!ov) {
      const app = document.getElementById('b8s-app'); if (!app) return;
      ov = document.createElement('div'); ov.id = 'b8s-bl'; app.appendChild(ov);
      ov.addEventListener('click', e => { if (e.target === ov) { BL.show = false; blRender(); return; } const el = e.target.closest('[data-a]'); if (el) blAction(el.dataset.a, el); });
    }
    const keep = ov.querySelector('.blp'), st = keep ? keep.scrollTop : 0;
    ov.innerHTML = blPanelHtml();
    const np = ov.querySelector('.blp'); if (np && st) np.scrollTop = st;
  }
  function blListOf(orders, tiktok, statusMap, now) {
    const H = orders.headers.map(h => String(h ?? '').normalize('NFC')), wb = orders.wbCol, rows = orders.rows;
    const stCol = colBy(H, /trạng thái/i), scCol = colBy(H, /người quét|nhân viên|shipper|người phát|giao nhận/i), dCol = colBy(H, /ngày cần phát/i);
    if (dCol < 0) return null;
    const out = [];
    for (const r of rows) {
      const back = daysBefore(r[dCol], now);
      if (back === null || back < BL_MIN) continue;
      if ((stCol >= 0 ? classify(r[stCol], statusMap) : 'other') === 'ok') continue;
      out.push([String(r[wb] ?? '').trim(), String(r[dCol] ?? '').trim(), back, stCol >= 0 ? String(r[stCol] ?? '').trim() : '', scCol >= 0 ? String(r[scCol] ?? '').trim() : '', isTikTok(r[wb], tiktok) ? 1 : 0]);
    }
    out.sort((a, b) => b[2] - a[2] || String(a[0]).localeCompare(String(b[0])));
    return out;
  }
  function makeBlStore(factory) {
    let dbp = null;
    const open = () => dbp || (dbp = new Promise((res, rej) => {
      const rq = factory.open('b8s-bl', 1);
      rq.onupgradeneeded = () => { const db = rq.result; if (!db.objectStoreNames.contains('off')) db.createObjectStore('off', { keyPath: 'office' }); };
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => { dbp = null; rej(rq.error || new Error('Không mở được kho danh sách backlog')); };
      rq.onblocked = () => { dbp = null; rej(new Error('Kho danh sách backlog đang bị tab khác giữ')); };
    }));
    const tx = async (mode, fn) => {
      const db = await open();
      return new Promise((res, rej) => {
        const t = db.transaction('off', mode); let out; const r = fn(t.objectStore('off'));
        if (r) r.onsuccess = () => { out = r.result; };
        t.oncomplete = () => res(out); t.onerror = t.onabort = () => rej(t.error || new Error('Lỗi ghi kho danh sách backlog'));
      });
    };
    return { all: () => tx('readonly', st => st.getAll()), put: r => tx('readwrite', st => { st.put(r); }) };
  }
  const blStore = (() => { try { return makeBlStore(indexedDB); } catch { return null; } })();
  let blEnsuring = false;
  function blEnsure() {
    if (BL.ready || blEnsuring) return;
    blEnsuring = true;
    (async () => {
      try { for (const r of (blStore ? await blStore.all() : []) || []) if (r && r.office && Array.isArray(r.items)) BL.map.set(r.office, r); }
      catch (e) { BL.err = `Không đọc được kho danh sách backlog: ${e.message}`; }
      BL.ready = true; blEnsuring = false;
      blRender();
    })();
  }
  const blAt = ms => { const d = new Date(ms); return Number.isFinite(+d) ? `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')} ${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}` : ''; };
  async function blLoadOne(office) {
    const { t } = await apiFetchOffice(office), now = new Date();
    const items = blListOf(t.orders, parseList(state.tiktok), state.statusMap, now);
    if (!items) throw new Error('đơn không có cột “Ngày cần phát”');
    const rec = { office, at: now.getTime(), items };
    if (blStore) { try { await blStore.put(rec); } catch (e) { BL.err = `Không lưu được danh sách ${office}: ${e.message}`; } }
    BL.map.set(office, rec);   // ghi đè danh sách cũ của bưu cục này
    return rec;
  }
  async function blLoad(list) {
    if (BL.run) return;
    BL.run = { i: 0, n: list.length, cur: '' }; BL.msg = ''; BL.err = ''; blRender();
    const bad = [];
    for (let i = 0; i < list.length; i++) {
      BL.run.i = i; BL.run.cur = list[i]; blRender();
      try { await blLoadOne(list[i]); } catch (e) { bad.push(`${list[i]}: ${e && e.message || e}`); if (isAuthError(e)) break; }
    }
    BL.run = null;
    BL.msg = bad.length ? '' : `Đã tải xong ${list.length} bưu cục.`;
    if (bad.length) BL.err = `Chưa tải được: ${bad.join('; ')}`;
    blRender();
  }
  const blCodes = rec => rec.items.map(x => x[0]).join('\n');
  const blCsv = rec => { const A = blHasAddr(rec); return '﻿' + [['Mã vận đơn', 'Ngày cần phát', 'Trễ (ngày)', 'Trạng thái', 'Người quét mới nhất', 'TikTok'].concat(A ? ['Địa chỉ giao'] : [])].concat(rec.items.map(x => [x[0], x[1], x[2], x[3], x[4], x[5] ? 'Có' : ''].concat(A ? [blAddrText(BLA.map.get(x[0]))] : []))).map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(',')).join('\r\n'); };
  function blPanelHtml() {
    blEnsure();
    const offs = anmOffices().filter(o => anmSel().has(o));
    if (!offs.length) return '';
    const run = BL.run, rowH = o => {
      const rec = BL.map.get(o), n = rec ? rec.items.length : 0;
      const busy = !!run, cur = run && run.cur === o;
      return `<div class="blr${BL.open === o ? ' on' : ''}"><b class="blo">${esc(o)}</b><span class="bln">${rec ? `<b>${fmt(n)}</b> đơn <small>· tải lúc ${blAt(rec.at)}</small>` : '<small>chưa tải</small>'}</span>
        <span class="blb"><button class="btn" data-a="blOne" data-o="${esc(o)}"${busy || BLA.run ? ' disabled' : ''}>${cur ? 'Đang tải…' : rec ? 'Tải lại' : 'Tải'}</button>${rec && rec.items.length ? `<button class="btn" data-a="blAddr" data-o="${esc(o)}"${busy || BLA.run ? ' disabled' : ''} title="Lấy địa chỉ giao của các đơn trong danh sách. Chỉ giữ trong phiên, không lưu.">${BLA.run && BLA.run.o === o ? `Địa chỉ ${BLA.run.i}/${BLA.run.n}…` : blHasAddr(rec) ? 'Tải lại địa chỉ' : 'Lấy địa chỉ'}</button>` : ''}${rec ? `<button class="btn" data-a="blView" data-o="${esc(o)}">${BL.open === o ? 'Ẩn' : 'Xem'}</button><button class="btn" data-a="blCopy" data-o="${esc(o)}">${BL.copied === o ? '✔ Đã copy' : 'Copy mã'}</button><button class="btn" data-a="blCsv" data-o="${esc(o)}">CSV</button>` : ''}</span></div>`;
    };
    const rec = BL.open ? BL.map.get(BL.open) : null;
    const hasA = blHasAddr(rec);
    const tbl = rec ? `<div class="blt"><table><thead><tr><th>#</th><th>Mã vận đơn</th><th>Ngày cần phát</th><th>Trễ</th><th>Trạng thái</th><th>Người quét mới nhất</th><th>TikTok</th>${hasA ? '<th>Địa chỉ giao</th>' : ''}</tr></thead><tbody>${rec.items.map((x, i) => `<tr><td>${i + 1}</td><td>${esc(x[0])}</td><td>${esc(x[1])}</td><td><b>${x[2]}</b> ngày</td><td>${esc(x[3])}</td><td>${esc(x[4])}</td><td>${x[5] ? '✔' : ''}</td>${hasA ? `<td class="bla">${BLA.map.has(x[0]) ? esc(blAddrText(BLA.map.get(x[0]))) : '…'}</td>` : ''}</tr>`).join('') || `<tr><td colspan="${hasA ? 8 : 7}">Không có đơn backlog.</td></tr>`}</tbody></table></div>` : '';
    return `<div class="b8s-panel blp"><div class="anm-hd" style="margin:0 0 6px"><h3 style="margin:0">Danh sách đơn backlog từng bưu cục</h3><span style="display:flex;gap:6px"><button class="btn" data-a="blClose">✕ Đóng</button><button class="btn solid" data-a="blAll"${run ? ' disabled' : ''}>${run ? `Đang tải ${run.i + 1}/${run.n}…` : `Tải tất cả (${offs.length} bưu cục)`}</button></span></div>
      <div class="muted" style="font-size:12px;margin-bottom:6px">Đơn backlog 5 ngày: ngày cần phát cách hôm nay từ 5 ngày trở lên và chưa giao thành công. Chỉ tải khi bấm. Tải lại bưu cục nào thì danh sách cũ của bưu cục đó bị ghi đè.</div>
      ${BL.err ? `<div class="al a">${esc(BL.err)}</div>` : ''}${BLA.err ? `<div class="al a">${esc(BLA.err)}</div>` : ''}${BL.msg ? `<div class="al ok">✔ ${esc(BL.msg)}</div>` : ''}
      ${offs.map(rowH).join('')}${tbl}</div>`;
  }
  function blAction(act, el) {
    const o = el.dataset.o, rec = o ? BL.map.get(o) : null;
    if (act === 'blAll') blLoad(anmOffices().filter(x => anmSel().has(x))).catch(er => { BL.run = null; BL.err = er.message; blRender(); });
    else if (act === 'blOne') blLoad([o]).catch(er => { BL.run = null; BL.err = er.message; blRender(); });
    else if (act === 'blAddr') blAddrFetch(o).catch(er => { BLA.run = null; BLA.err = er.message; blRender(); });
    else if (act === 'blClose') { BL.show = false; blRender(); }
    else if (act === 'blView') { BL.open = BL.open === o ? '' : o; blRender(); }
    else if (act === 'blCopy' && rec) copyText(blCodes(rec)).then(ok => { if (ok) { BL.copied = o; setTimeout(() => { BL.copied = ''; blRender(); }, 2600); } blRender(); });
    else if (act === 'blCsv' && rec) { const d = new Date(rec.at); downloadText(`Backlog_${o.replace(/\s+/g, '_')}_${String(d.getDate()).padStart(2, '0')}${String(d.getMonth() + 1).padStart(2, '0')}${d.getFullYear()}.csv`, blCsv(rec), 'text/csv;charset=utf-8'); }
  }

  function renderThresholds() {
    const box = $('#b8s-th');
    if (!box) return;
    const row = (k, lab, fields) => `<tr><td class="l">${esc(lab)}</td>${fields.map(([f, l]) => `<td>${esc(l)} <input type="number" step="any" style="width:70px" data-th="${k}.${f}" value="${Object.assign({}, ANM_TH[k], state.th[k])[f]}"></td>`).join('')}</tr>`;
    box.innerHTML = `<table><tbody>${ANM_METRICS.filter(m => ANM_TH[m.k]).map(m => ANM_TH[m.k].lo !== undefined
      ? row(m.k, m.label + ' (%)', [['lo', 'Đỏ nếu dưới'], ['hi', 'Xanh từ']])
      : row(m.k, m.label, [['ok', 'Xanh nếu ≤'], ['warn', 'Đỏ nếu >']])).join('')}</tbody></table><div class="row" style="margin-top:8px"><button class="b8sbtn solid" id="b8s-th-reset">Đặt lại mặc định</button></div>`;
  }

  /* ------------------------------------------------------------------ *
   * Lưu cấu hình
   * ------------------------------------------------------------------ */
  function loadPrefs() {
    Object.assign(state.wx, WX_DEFAULT);
    state.dm = dmRulesOf(null);
    try {
      const p = JSON.parse(localStorage.getItem('b8s_prefs') || '{}');
      if (p.wx && typeof p.wx === 'object') {
        if (typeof p.wx.on === 'boolean') state.wx.on = p.wx.on;
        if (Number.isFinite(p.wx.lat) && Math.abs(p.wx.lat) <= 90) state.wx.lat = p.wx.lat;
        if (Number.isFinite(p.wx.lon) && Math.abs(p.wx.lon) <= 180) state.wx.lon = p.wx.lon;
      }
      if (typeof p.tiktok === 'string') state.tiktok = p.tiktok;
      if (p.dm && typeof p.dm === 'object') state.dm = dmRulesOf(p.dm);
      if (p.th && typeof p.th === 'object') state.th = p.th;
      if (p.anm && typeof p.anm === 'object') {
        const a = state.anm;
        if (Array.isArray(p.anm.sel)) a.sel = new Set(p.anm.sel.filter(x => typeof x === 'string'));
        if (CMP_MODES.includes(p.anm.cmp)) a.cmp = p.anm.cmp;
        if (RANK_KEYS.includes(p.anm.rankKey)) a.rankKey = p.anm.rankKey;
        if (RANK_KEYS.includes(p.anm.trendKey)) a.trendKey = p.anm.trendKey;
      }
      if (typeof p.fast === 'boolean') state.fast = p.fast;
      if (p.statusMap && typeof p.statusMap === 'object') state.statusMap = p.statusMap;
      if (Array.isArray(p.recent)) state.recent = p.recent.filter(x => typeof x === 'string').slice(0, 20);
    } catch { /* bỏ qua */ }
  }
  function savePrefs() {
    try {
      localStorage.setItem('b8s_prefs', JSON.stringify({ th: state.th, anm: { sel: state.anm.sel ? [...state.anm.sel] : null, cmp: state.anm.cmp, rankKey: state.anm.rankKey, trendKey: state.anm.trendKey }, tiktok: state.tiktok, dm: state.dm, fast: state.fast, statusMap: state.statusMap, recent: state.recent, wx: { on: state.wx.on, lat: state.wx.lat, lon: state.wx.lon } }));
    } catch { /* bỏ qua */ }
  }

  /* ------------------------------------------------------------------ *
   * Giao diện
   * ------------------------------------------------------------------ */
  function setBusy(b) {
    document.querySelectorAll('#b8s-app .b8sbtn[data-act]').forEach(btn => { btn.disabled = b; btn.style.opacity = b ? .6 : 1; });
  }

  function getCalc() {
    if (!state.orders) return null;
    const key = [state.orders.at.getTime(), state.emp ? state.emp.at.getTime() : 0, state.tiktok, JSON.stringify(state.statusMap), JSON.stringify(state.dm), new Date().toDateString()].join('|');
    if (state.calcKey !== key) {
      state.calc = computeDashboard({ orders: state.orders, emp: state.emp, tiktok: parseList(state.tiktok), statusMap: state.statusMap, now: new Date(), rules: state.dm, skip: shipSkipFn(state.office) });
      state.calcKey = key;
    }
    return state.calc;
  }

  /* ------------------------------------------------------------------ *
   * Bảng Shipper: cột "Tỉ lệ giao Tuần" và "Tỉ lệ giao Tháng" (chỉ tải khi bấm nút)
   * Nguồn: danh sách đơn Kanban của từng ngày (mỗi đơn có mã ship và trạng thái ký nhận).
   * Tuần = tuần trước (T2 đến CN); tháng = từ mùng 1 đến hôm qua. Tính cả đơn TikTok.
   * Tỉ lệ = tổng đơn đã ký nhận / tổng đơn của ship, cộng dồn các ngày. Chỉ lưu số tổng theo ship và ngày.
   * ------------------------------------------------------------------ */
  const SH_KEEP = 75, SH_WORKERS = 3;
  // Các ngày của kỳ: w = tuần trước, m = từ mùng 1 đến hôm qua
  function shDays(p, today) {
    if (p === 'w') {
      const dow = new Date(keyToUTC(today)).getUTCDay(), w0 = addDays(addDays(today, -((dow + 6) % 7)), -7);
      return Array.from({ length: 7 }, (_, i) => addDays(w0, i));
    }
    const m0 = `${today.slice(0, 8)}01`, n = Math.round((keyToUTC(today) - keyToUTC(m0)) / 86400000);
    return Array.from({ length: n }, (_, i) => addDays(m0, i));
  }
  // Gộp danh sách đơn một ngày thành { mã ship: [tổng đơn, đã ký nhận] }
  function shAggregate(items) {
    const seen = new Set(), out = {};
    for (const r of items || []) {
      if (!r || typeof r !== 'object') continue;
      const b = r.billCode ? String(r.billCode) : '';
      if (b) { if (seen.has(b)) continue; seen.add(b); }
      const code = shipKeys(r.smallManCode).code;
      if (!code) continue;
      const x = out[code] || (out[code] = [0, 0]);
      x[0]++; if (Number(r.isSign) === 1) x[1]++;
    }
    return out;
  }
  // Cộng các ngày của một kỳ cho từng mã ship; days thiếu thì trả về thiếu
  function shPeriod(map, office, days) {
    const by = {}; let have = 0;
    for (const d of days) {
      const r = map.get(`${office}|${d}`); if (!r) continue;
      have++;
      for (const [c, v] of Object.entries(r.s)) { const x = by[c] || (by[c] = [0, 0]); x[0] += v[0]; x[1] += v[1]; }
    }
    return { by, have, need: days.length };
  }
  const shPct = (ok, tot) => tot > 0 ? `${(ok / tot * 100).toFixed(2).replace('.', ',')}%` : '—';
  function makeShStore(factory) {
    let dbp = null;
    const open = () => dbp || (dbp = new Promise((res, rej) => {
      const rq = factory.open('b8s-sh', 1);
      rq.onupgradeneeded = () => { const db = rq.result; if (!db.objectStoreNames.contains('days')) db.createObjectStore('days', { keyPath: 'id' }); };
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => { dbp = null; rej(rq.error || new Error('Không mở được kho số liệu ship')); };
      rq.onblocked = () => { dbp = null; rej(new Error('Kho số liệu ship đang bị tab khác giữ')); };
    }));
    const tx = async (mode, fn) => {
      const db = await open();
      return new Promise((res, rej) => {
        const t = db.transaction('days', mode); let out; const r = fn(t.objectStore('days'));
        if (r) r.onsuccess = () => { out = r.result; };
        t.oncomplete = () => res(out); t.onerror = t.onabort = () => rej(t.error || new Error('Lỗi ghi kho số liệu ship'));
      });
    };
    return { all: () => tx('readonly', st => st.getAll()), put: r => tx('readwrite', st => { st.put(r); }), delMany: ids => tx('readwrite', st => { ids.forEach(i => st.delete(i)); }) };
  }
  const shStore = (() => { try { return makeShStore(indexedDB); } catch { return null; } })();
  const shOn = () => !!state.office && !isFS(state.office);
  async function shLoadStore() {
    const S = state.sh; if (S.ready) return;
    if (shStore) {
      try {
        const lim = addDays(dateKey(new Date()), -SH_KEEP), dead = [];
        for (const r of (await shStore.all()) || []) {
          if (!r || !r.id || !r.s || typeof r.s !== 'object') { if (r && r.id) dead.push(r.id); continue; }
          if (r.date < lim) dead.push(r.id); else S.map.set(r.id, r);
        }
        if (dead.length) await shStore.delMany(dead);
      } catch (e) { S.err0 = `Không đọc được kho số liệu ship: ${e.message}`; }
    }
    S.ready = true;
  }
  let shRedrawT = 0;
  function shRedraw(now) {
    if (state.view !== 'dash') return;
    if (now) { clearTimeout(shRedrawT); shRedrawT = 0; renderDashboard(); return; }
    if (!shRedrawT) shRedrawT = setTimeout(() => { shRedrawT = 0; if (state.view === 'dash') renderDashboard(); }, 300);
  }
  // Bấm nút tải: chỉ gọi dữ liệu của bưu cục đang xem, chỉ cho kỳ được bấm
  async function shLoad(p) {
    const S = state.sh, office = state.office;
    if (!office || isFS(office) || S.run[p]) return;
    if (!state.fast) { S.msg = 'Đang tắt “Tải nhanh bằng API” trong Cài đặt nên không tải được số liệu ship.'; shRedraw(true); return; }
    await shLoadStore();
    const today = dateKey(new Date()), days = shDays(p, today);
    S.msg = '';
    if (!days.length) { S.res[`${office}|${p}`] = { at: Date.now(), days, empty: true }; shRedraw(true); return; }
    const force = !!(S.res[`${office}|${p}`] && S.res[`${office}|${p}`].at);
    const jobs = days.filter(d => force ? true : !S.map.has(`${office}|${d}`));
    S.run[p] = true; S.prog[p] = `Đang tải 0/${jobs.length}…`; shRedraw(true);
    let i = 0, done = 0, stop = false;
    const worker = async () => {
      while (!stop && i < jobs.length) {
        const d = jobs[i++];
        try {
          const site = await resolveSite(office), day = await kbFetchDay(site.code, d, true);
          const rec = { id: `${office}|${d}`, office, date: d, at: Date.now(), s: shAggregate(day.items), n: day.items.length, total: day.total };
          S.map.set(rec.id, rec);
          if (shStore) { try { await shStore.put(rec); } catch (e) { S.err0 = `Không lưu được số liệu ship: ${e.message}`; } }
        } catch (e) {
          stop = true; S.msg = isAuthError(e) ? 'Cần đăng nhập lại trang 800BEST để lấy số liệu ship.' : `Không tải được ngày ${fmtDM(d)}: ${e.message}`;
        }
        done++; S.prog[p] = `Đang tải ${done}/${jobs.length}…`; shRedraw();
      }
    };
    try { await Promise.all(Array.from({ length: Math.min(SH_WORKERS, jobs.length) }, worker)); }
    finally {
      S.run[p] = false; S.prog[p] = '';
      if (days.every(d => S.map.has(`${office}|${d}`))) S.res[`${office}|${p}`] = { at: Date.now(), days };
      shRedraw(true);
    }
  }
  // Ô tỉ lệ của một ship (null = chưa bấm tải)
  function shCell(r, p) {
    const S = state.sh, res = S.res[`${state.office}|${p}`];
    if (!res) return '<td class="sh-e"></td>';
    if (res.empty) return '<td class="sh-na">chưa có ngày nào</td>';
    const c = shipKeys(r.name).code, v = c ? shPeriod(S.map, state.office, res.days).by[c] : null;
    if (!v || !v[0]) return '<td class="sh-na">không có số liệu</td>';
    return `<td class="${rateClass(v[1] / v[0] * 100, 'high')}">${esc(shPct(v[1], v[0]))}<small>${fmt(v[1])}/${fmt(v[0])}</small></td>`;
  }
  function shTotalCell(rows, p) {
    const S = state.sh, res = S.res[`${state.office}|${p}`];
    if (!res) return '<td class="sh-e"></td>';
    if (res.empty) return '<td class="sh-na">—</td>';
    const by = shPeriod(S.map, state.office, res.days).by; let ok = 0, tot = 0;
    for (const r of rows) { const c = shipKeys(r.name).code, v = c ? by[c] : null; if (v) { tot += v[0]; ok += v[1]; } }
    return tot ? `<td class="${rateClass(ok / tot * 100, 'high')}">${esc(shPct(ok, tot))}<small>${fmt(ok)}/${fmt(tot)}</small></td>` : '<td class="sh-na">—</td>';
  }
  function shHead(p, label) {
    const S = state.sh, res = S.res[`${state.office}|${p}`], run = S.run[p];
    const t = run ? S.prog[p] : res ? `Tải lại · ${pad(new Date(res.at).getHours())}:${pad(new Date(res.at).getMinutes())}` : 'Tải dữ liệu';
    return `<th class="g-sh">${label} <button class="sh-ld" data-sh="${p}" ${run ? 'disabled' : ''}>${esc(t)}</button></th>`;
  }

  // Nhãn rút gọn cho dòng phụ của thẻ chỉ số (đầy đủ nằm trong chú thích khi rê chuột)
  const CARD_ABBR = { 'Giao thành công': 'GTC', 'Tổng đơn mới hôm nay': 'đơn mới', 'Chờ duyệt': 'duyệt', 'Giao thất bại': 'thất bại', 'Nhập kho chưa quét phát': 'kho' };
  function rateClass(v, good) {
    if (typeof v !== 'number' || !Number.isFinite(v)) return '';
    if (good === 'low') return v <= 10 ? 'r-ok' : v <= 20 ? 'r-mid' : 'r-bad';
    return v >= 75 ? 'r-ok' : v >= 60 ? 'r-mid' : 'r-bad';
  }

  // v3.31: xem lại ngày cũ ở tab Tổng quan, dựng từ bản ghi đã lưu của bưu cục đang chọn
  function recToCalc(r) {
    const kb = r.rb === 'kb';
    const card = a => ({
      pod: ratio(a.newOk, a.newTotal), newOk: a.newOk, newTotal: a.newTotal, newUns: null, newScan: null,
      total: a.total, unscanned: kb ? null : a.unscanned, scanned: kb ? null : a.total - a.unscanned, ok: a.ok,
      scanRate: kb ? null : ratio(a.total - a.unscanned, a.total),
      perNV: r.nWork > 0 ? a.total / r.nWork : null,
      gtcTrip: kb ? null : ratio(a.ok, a.total - a.unscanned), gtcTotal: ratio(a.ok, a.total)
    });
    const shippers = r.ships.map(x => ({ name: x.n, matched: true, scanned: kb ? null : x.sc, ok: x.ok, ttTotal: kb ? null : x.tt, ttOk: kb ? null : x.tk,
      rateTrip: kb ? null : ratio(x.ok, x.sc), rateTT: kb ? null : ratio(x.tk, x.tt) }));
    const sum = k => shippers.reduce((a, x) => a + (Number.isFinite(x[k]) ? x[k] : 0), 0);
    const totals = { name: 'TỔNG', scanned: kb ? null : sum('scanned'), ok: sum('ok'), ttTotal: kb ? null : sum('ttTotal'), ttOk: kb ? null : sum('ttOk') };
    totals.rateTrip = kb ? null : ratio(totals.ok, totals.scanned); totals.rateTT = kb ? null : ratio(totals.ttOk, totals.ttTotal);
    const all = card(r.all), tt = card(r.tt), sme = { unscanned: kb ? null : r.all.unscanned - r.tt.unscanned };
    const at = r.savedAt ? new Date(r.savedAt) : null;
    const notes = [`Đang xem bản đã lưu ngày ${fmtDM(r.date)}${at && !isNaN(at) ? ` (lưu lúc ${pad(at.getHours())}:${pad(at.getMinutes())})` : ''}. Bản lưu chỉ có số tổng và số theo shipper, không có danh sách đơn.`];
    if (kb) notes.push('Bản này được dựng lại từ danh sách đơn của Kanban nên không có số quét giao (các ô liên quan để trống).');
    return { cards: { all, tt, sme }, extra: { backlog: r.backlog, odrTT: r.odrTT, odr2: r.odr2 }, firstOk: r.firstOk, hasTT: tt.total > 0,
      nWork: r.nWork, workersEstimated: r.workersEstimated, shippers, totals, notes };
  }
  function dashDayBar(today, past) {
    const day = state.dashDay || today;
    return `<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:0 0 4px;color:#fff;font-size:12px;line-height:1.2"><label>Ngày <input type="date" id="b8s-dday" value="${esc(day)}" max="${today}"></label>${past ? '<span class="b8sbtn" id="b8s-dtoday">Về hôm nay</span><span class="muted">Đang xem số đã lưu, không phải số mới.</span>' : ''}</div>`;
  }
  function bindDashDay(box) {
    const i = box.querySelector('#b8s-dday'), b = box.querySelector('#b8s-dtoday');
    if (i) i.onchange = () => { const v = i.value, t = dateKey(new Date()); state.dashDay = v && v < t ? v : ''; renderDashboard(); };
    if (b) b.onclick = () => { state.dashDay = ''; renderDashboard(); };
  }
  function officeListPanelHtml() {
    const L = state.ofList;
    const btn = `<span class="btn" data-act="officeList">${L.busy ? '⟳ Đang tải danh sách bưu cục…' : '🔍 Tải danh sách bưu cục'}</span>`;
    const listTxt = OFFICES.length ? `<div class="muted" style="margin-top:4px">Đang quản lý <b>${OFFICES.length}</b> bưu cục: ${esc(OFFICES.join(', '))}</div>` : '<div class="muted" style="margin-top:4px">Chưa có bưu cục nào — bấm nút trên để tool hỏi hệ thống xem tài khoản đang đăng nhập quản lý bưu cục nào.</div>';
    const note = L.note ? `<div class="al ${L.warn ? 'a' : 'ok'}" style="margin-top:4px">${esc(L.note)}</div>` : '';
    return `<div class="b8s-panel" style="padding:10px 14px"><div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap"><b>Bưu cục đang quản lý</b>${btn}<span class="muted">Danh sách này dùng chung cho cả tab ANM. Bấm lại để thay mới hoàn toàn theo tài khoản V9 đang đăng nhập.</span></div>${listTxt}${note}</div>`;
  }
  function renderDashboard() {
    const box = $('#b8s-dash');
    if (!box) return;
    const ofPanel = officeListPanelHtml();
    const today0 = dateKey(new Date());
    if (state.dashDay && state.dashDay >= today0) state.dashDay = '';
    const past = !!state.dashDay;
    let c = null;
    if (past) {
      const rec = state.office ? state.hist.get(`${state.office}|${state.dashDay}`) || [...state.hist.values()].find(r => r.date === state.dashDay && sameOffice(r.office, state.office)) : null;
      if (!rec) { box.innerHTML = ofPanel + dashDayBar(today0, true) + `<div class="b8s-empty"><b>Không có bản lưu ngày ${esc(fmtDM(state.dashDay))}${state.office ? ' của ' + esc(state.office) : ''}</b><br>${state.office ? 'Chọn ngày khác hoặc “Về hôm nay”.' : 'Chọn bưu cục ở thanh phía trên (bấm “Chuyển bưu cục”) để xem ngày cũ của bưu cục đó.'}</div>`; bindDashDay(box); return; }
      c = recToCalc(rec);
    } else c = getCalc();
    const dayBar = dashDayBar(today0, past);
    if (!c) {
      box.innerHTML = ofPanel + dayBar + '<div class="b8s-empty"><b>Chưa có dữ liệu</b><br>Chọn bưu cục ở thanh phía trên rồi bấm “Chuyển bưu cục”. Script sẽ tự đọc toàn bộ số liệu (tab Tiến độ nhân viên, Chi tiết phát kiện) rồi quay về tab bạn đang xem.</div>';
      bindDashDay(box); return;
    }
    const cl = cardList(c);
    const card = ([label, v, kind, color, sub]) => (v === null || v === undefined || !Number.isFinite(v)) ? ''
      : `<div class="b8s-card" title="${esc(Array.isArray(sub) ? sub.map(([k, n]) => `${k}: ${fmt(n)}`).join(' · ') : label)}"><span>${esc(label)}</span><b class="c-${color}">${esc(fmtVal(v, kind))}</b><small>${
          Array.isArray(sub) ? sub.map(([k, n]) => `${esc(CARD_ABBR[k] || k)} <i>${esc(fmt(n))}</i>`).join(' · ') : '&nbsp;'}</small></div>`;
    const cols = shipperCols(c);
    const gh = cols.filter(x => x.group === 'gh').length, tt = cols.filter(x => x.group === 'tt').length;
    const rows = sortRows(c.shippers, state.sort.key, state.sort.dir);
    const th = (x, extra = '') => `<th class="${x.k === 'name' ? 'l ' : ''}g-${x.group || 'nav'} ${extra}" data-k="${x.k}" title="${esc(x.tip || '')}">${esc(x.label)}${state.sort.key === x.k ? (state.sort.dir > 0 ? ' ▲' : ' ▼') : ''}</th>`;
    const td = (x, r) => {
      const v = r[x.k];
      if (x.kind === 'text') return `<td class="l">${esc(v)}</td>`;
      if (x.kind === 'pct') return `<td class="${rateClass(v, x.good)}">${esc(fmtVal(v, 'pct'))}</td>`;
      return `<td>${esc(fmtVal(v, 'int'))}</td>`;
    };
    const sho = !past && shOn(), ci = cols.findIndex(x => x.k === 'name') + 1;
    const head2 = `<tr><th class="g-nav">#</th>${cols.map((x, i) => th(x) + (sho && i === ci - 1 ? (state.sh.hide ? '<th class="g-sh"><button class="sh-ld" data-shtog="1" title="Hiện lại hai cột tuần và tháng">Hiện</button></th>' : shHead('w', 'Tỉ lệ giao Tuần') + shHead('m', 'Tỉ lệ giao Tháng')) : '')).join('')}</tr>`;
    const head1 = `<tr><th class="g-nav" colspan="2">Shipper đi làm: ${fmt(c.nWork)}${c.workersEstimated ? ' (ước tính)' : ''}</th>${sho ? (state.sh.hide ? '<th class="g-sh">Tỉ lệ giao</th>' : '<th class="g-sh" colspan="2">Tỉ lệ giao thành công <button class="sh-ld" data-shtog="1" title="Ẩn hai cột tuần và tháng (số đã tải vẫn giữ)">Ẩn</button></th>') : ''}${gh ? `<th class="g-gh" colspan="${gh}">Chỉ số giao hàng</th>` : ''}${tt ? `<th class="g-tt" colspan="${tt}">Chỉ số TikTok Shop</th>` : ''}</tr>`;
    const body = rows.map((r, i) => `<tr><td class="rk${i < 3 ? i + 1 : ''}">${i + 1}</td>${cols.map((x, k) => td(x, r) + (sho && k === ci - 1 ? (state.sh.hide ? '<td class="sh-e"></td>' : shCell(r, 'w') + shCell(r, 'm')) : '')).join('')}</tr>`).join('');
    const total = `<tr class="total"><td></td>${cols.map((x, k) => td(x, c.totals) + (sho && k === ci - 1 ? (state.sh.hide ? '<td class="sh-e"></td>' : shTotalCell(rows, 'w') + shTotalCell(rows, 'm')) : '')).join('')}</tr>`;

    // Đối chiếu với trang gốc
    const notes = [...c.notes];
    if (!past && notes.some(t => /sai chức vụ/.test(t)) && state.emp && Array.isArray(state.emp.rows)) {   // chẩn đoán: cho thấy đúng giá trị “Chức vụ” đang đọc được để so với danh sách ở Cài đặt
      const di = (state.emp.headers || []).findIndex(h => /chức vụ/i.test(String(h))), seen = new Map();
      if (di >= 0) for (const r of state.emp.rows) { const v = String(r[di] ?? '').trim() || '(trống)'; seen.set(v, (seen.get(v) || 0) + 1); }
      if (seen.size) notes.push(`Chức vụ đọc được từ hệ thống: ${[...seen].map(([k, n]) => `“${k}” ×${n}`).join(' · ')}. Nếu đúng là ship đi làm, vào Cài đặt thêm đúng dòng này vào “Chức vụ được tính” (hoặc tắt lọc chức vụ).`);
    }
    if (sho) {
      const S = state.sh, rw = S.res[`${state.office}|w`], rm = S.res[`${state.office}|m`], tdy = dateKey(new Date());
      if ((rw || rm) && !S.hide) {
        const rg = r => r && r.days && r.days.length ? `${fmtDM(r.days[0])} đến ${fmtDM(r.days[r.days.length - 1])}` : '';
        notes.push(`Tỉ lệ giao thành công${rw ? ` · tuần trước ${rg(rw)}` : ''}${rm ? ` · tháng ${rm.empty ? 'chưa có ngày nào (không gồm hôm nay)' : rg(rm) + ' (không gồm hôm nay)'}` : ''} · đã ký nhận / tổng đơn của ship, có cả đơn TikTok.`);
      }
      if (S.msg) notes.push(`⚠ ${S.msg}`);
      if (S.err0) notes.push(`⚠ ${S.err0}`);
    }
    const ov = past ? null : state.ov;
    if (ov) {
      const chk = (label, a, b) => {
        if (!Number.isFinite(a)) return;
        notes.push(a === b ? `✔ ${label}: ${fmt(b)} khớp trang gốc.` : `⚠ ${label}: tự tính ${fmt(b)} nhưng trang gốc ghi ${fmt(a)}. Kiểm tra lại phân loại trạng thái ở Cài đặt.`);
      };
      chk('Tổng đơn cần phát', ov.need, c.cards.all.total);
      chk('Đơn nhập kho chưa quét phát', ov.unscanned, c.cards.all.unscanned);
    }

    const warn = (!past && state.pending) ? `<div class="b8s-warn">⚠ Chưa lưu lịch sử: ${esc(state.pending.reason)} <button class="b8sbtn solid" data-act="force-save">Vẫn lưu</button></div>` : '';
    box.innerHTML = `${ofPanel}${dayBar}${warn}
      <div class="b8s-band"><div class="bl">📦 Giao thường</div>${cl.main.map(card).join('')}</div>
      ${c.hasTT ? `<div class="b8s-band"><div class="bl">🎵 TikTok Shop</div>${cl.tts.map(card).join('')}</div>` : ''}
      <div class="b8s-tw"><table><thead>${head1}${head2}</thead><tbody>${body}${total}</tbody></table></div>
      ${notes.length ? `<div class="b8s-notes">${notes.map(n => `<div>${esc(n)}</div>`).join('')}</div>` : ''}`;
    bindDashDay(box);
  }

  function renderSettings() {
    renderOffices();
    const box = $('#b8s-map');
    if (!box) return;
    if (!state.orders) { box.innerHTML = '<div class="muted">Chưa có dữ liệu đơn. Chọn bưu cục rồi bấm “Chuyển bưu cục” để liệt kê các trạng thái.</div>'; box.dataset.key = ''; return; }
    const stCol = colBy(state.orders.headers, /trạng thái/i);
    if (stCol < 0) { box.innerHTML = '<div class="muted bad">Không thấy cột “Trạng thái” trong bảng đơn.</div>'; return; }
    const list = statusCounts(state.orders.rows, stCol);
    const sig = JSON.stringify(list) + JSON.stringify(state.statusMap);
    if (box.dataset.key === sig) return;
    box.dataset.key = sig;
    const opt = (v, cur, l) => `<option value="${v}" ${v === cur ? 'selected' : ''}>${l}</option>`;
    box.innerHTML = `<table><thead><tr><th class="l">Trạng thái trong bảng</th><th>Số đơn</th><th class="l">Tính là</th></tr></thead><tbody>${
      list.map(([s, n]) => {
        const cur = classify(s, state.statusMap);
        return `<tr><td class="l">${esc(s)}</td><td>${fmt(n)}</td><td class="l"><select data-status="${esc(s)}">${opt('ok', cur, 'Giao thành công')}${opt('unscanned', cur, 'Chưa quét giao (nhập kho)')}${opt('other', cur, 'Đã quét giao (đang giao, khác)')}</select></td></tr>`;
      }).join('')
    }</tbody></table>`;
  }

  function renderSnapInfo() {
    const box = $('#b8s-snapinfo'); if (!box) return;
    const S = state.snap, dates = [...new Set(S.list.map(r => r.date))].sort(), today = dateKey(new Date()), now = new Date();
    const todayTxt = SNAP_SLOTS.map(sl => { const st = slotState(today, sl, now, OFFICES, S.idx); return `${sl} ${st.st === 'ok' ? '✔' : st.st === 'part' ? `(${st.have}/${st.n})` : st.st === 'miss' ? 'lỡ' : st.st === 'due' ? 'đang chờ' : '–'}`; }).join(' · ');
    box.innerHTML = `Tự lưu số chốt ngày lúc <b>${SNAP_SLOTS.join(', ')}</b> (giờ máy, kể cả Chủ nhật) khi trang 800BEST đang mở; mở trễ trong ${SNAP_WINDOW_MIN} phút sau mốc thì chạy bù. Dùng để giữ số DM đi làm và giờ GTC đầu tiên của từng ngày. Giữ ${SNAP_KEEP_DAYS} ngày, cũ hơn tự xóa. Các bản chụp theo mốc giờ cũ vẫn còn trong trình duyệt nhưng không dùng nữa.<br>`
      + `Hôm nay: ${esc(todayTxt)}<br>Đang có ${S.list.length} bản chụp${dates.length ? ` (từ ${esc(fmtDMY(dates[0]))} đến ${esc(fmtDMY(dates[dates.length - 1]))})` : ''}.`
      + `${state.fast ? '' : '<br><b class="bad">Đang tắt “Tải nhanh bằng API” nên không tự lưu được.</b>'}${S.err ? `<br><b class="bad">${esc(S.err)}</b>` : ''}`
      + '<br>Mỗi bản chụp có thêm số liệu từng ship (tên, số quét, số giao thành công, giờ quét đầu) để phân tích nguyên nhân; không lưu mã vận đơn hay thông tin người nhận.'
      + '<br>Lưu ý: tool chỉ chạy được khi máy bật, trình duyệt mở và đã đăng nhập 800BEST. Số hôm qua và tuần trước để so sánh được lấy tự động từ Kanban (xem mục “Dữ liệu so sánh” bên dưới).';
  }

  function renderBkInfo() {
    const box = $('#b8s-bk-st'); if (!box) return;
    const bi = backupInfo();
    box.innerHTML = `Đang có ${state.hist.size} bản ghi ngày và ${state.snap.list.length} bản chụp theo mốc giờ.${bi.due ? ' <b class="bad">Đã hơn 7 ngày chưa sao lưu.</b>' : ''}`;
  }
  function renderUI() {
    if (!$('#b8s-app')) return;
    const title = $('#b8s-title'), sub = $('#b8s-sub');
    $('#b8s-app').classList.toggle('anm', state.view === 'anm' || state.view === 'pick');
    if (state.view === 'anm' || state.view === 'pick') { if (title) title.textContent = `📍 ${state.view === 'pick' ? 'Pickup-sale' : 'Thống kê'} Khu Vực${REGION ? ' ' + REGION : ''}`; }
    else if (title) title.textContent = state.office ? `📍 ${state.office}${state.province ? ' – ' + titleCase(state.province) : ''}` : '📊 Thống kê tổng quan';
    if (state.office && document.activeElement !== $('#b8s-office-in')) renderOfficeList();
    if (sub && state.view !== 'anm' && state.view !== 'pick') {
      let t = state.lastRefresh ? `Cập nhật lúc ${state.lastRefresh.toLocaleTimeString('vi-VN')} ngày ${state.lastRefresh.toLocaleDateString('vi-VN')}` : 'Chưa tải dữ liệu';
      if (state.lastSaved) t += ` · đã lưu lịch sử ${pad(state.lastSaved.getHours())}:${pad(state.lastSaved.getMinutes())}`;
      if (backupInfo().due) t += ' · ⚠ nên sao lưu lịch sử';
      sub.textContent = t;
    }
    ['dash', 'hist', 'anm', 'pick', 'set'].forEach(v => {
      const sec = $(`#b8s-v-${v}`); if (sec) sec.style.display = state.view === v ? 'block' : 'none';
    });
    document.querySelectorAll('#b8s-app .b8s-tab').forEach(t => t.classList.toggle('on', t.dataset.view === state.view));
    if (state.view === 'dash') renderDashboard();
    if (state.view === 'hist') renderHistory();
    if (state.view === 'anm') renderAnm();
    if (state.view === 'pick') renderPick();
    if (state.view === 'set') { renderSettings(); renderSnapInfo(); renderWxInfo(); renderBkInfo(); renderKbInfo(); }
    snapBadge();
  }

  /* ------------------------------------------------------------------ *
   * Xuất Excel & chẩn đoán
   * ------------------------------------------------------------------ */
  async function exportExcel() {
    if (state.busy) return;
    if (typeof XLSX === 'undefined') { log('Thư viện Excel chưa tải được. Kiểm tra mạng/CDN rồi tải lại trang.', true); return; }
    if (!state.orders && !state.emp && state.view !== 'hist' && state.view !== 'anm') { if (!(await updateAll())) return; }
    const c = getCalc();
    const now = new Date(), used = new Set(), wb = XLSX.utils.book_new();
    const addSheet = (name, aoa) => {
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = (aoa[0] || []).map((_, i) => ({ wch: Math.min(40, Math.max(10, ...aoa.slice(0, 200).map(r => String(r[i] ?? '').length + 2))) }));
      XLSX.utils.book_append_sheet(wb, ws, sheetName(name, used));
    };
    const val = (v, kind) => (kind === 'text') ? String(v ?? '') : (typeof v === 'number' && Number.isFinite(v)) ? (kind === 'int' ? v : round1(v)) : '';

    if (c) {
      const cl = cardList(c);
      const ov = [['Nhóm', 'Chỉ tiêu', 'Giá trị']];
      const push = (g, list) => list.forEach(([label, v, kind, , sub]) => {
        if (v === null || v === undefined || !Number.isFinite(v)) return;
        ov.push([g, label + (kind === 'pct' ? ' (%)' : ''), kind === 'int' ? v : round1(v)]);
        if (Array.isArray(sub)) sub.forEach(([k, n]) => ov.push([g, `${label} – ${k}`, n]));
      });
      push('Giao thường', cl.main);
      if (c.hasTT) push('TikTok Shop', cl.tts);
      addSheet('Tong quan', ov);

      const cols = shipperCols(c);
      const rows = sortRows(c.shippers, state.sort.key, state.sort.dir);
      addSheet('Shipper', [
        cols.map(x => x.label),
        ...rows.map(r => cols.map(x => val(r[x.k], x.kind))),
        cols.map(x => val(c.totals[x.k], x.kind))
      ]);
    }

    if (state.view === 'anm') addAnmSheets(addSheet);
    addHistorySheets(addSheet, now);

    const notes = [['Nội dung'], [`Bưu cục: ${state.office || 'Bưu cục đang chọn'}${state.province ? ' – ' + titleCase(state.province) : ''}`], [`Thời điểm xuất: ${now.toLocaleString('vi-VN')}`]];
    if (state.lastRefresh) notes.push([`Dữ liệu đọc lúc: ${state.lastRefresh.toLocaleString('vi-VN')}`]);
    notes.push([`Mã vận đơn TikTok Shop: ${state.tiktok}`]);
    notes.push(['POD2400 = đơn mới hôm nay (Ngày cần phát hàng = hôm nay) giao thành công / tổng đơn mới hôm nay, tính cả đơn chưa quét giao. Backlog 3 ngày = đơn có Ngày cần phát cách hôm nay từ 3 ngày trở lên và chưa giao thành công. TikTok đến hạn ODR = đơn TikTok có Thời hạn giao hàng nội bộ là hôm nay.']);
    notes.push(['GTC chuyến = đơn giao thành công / đơn đã quét giao; GTC Total = đơn giao thành công / tổng đơn cần phát. Tổng đơn TikTok ở bảng shipper = đơn TikTok đã quét giao.']);
    notes.push(['Tỷ lệ quét giao = (tổng đơn − đơn nhập kho chưa quét) / tổng đơn. TB đơn/NV = tổng đơn / số shipper đi làm hôm nay.']);
    if (c) c.notes.forEach(t => notes.push([t]));
    notes.push([state.includePII ? 'Có kèm thông tin người nhận.' : 'Không kèm thông tin người nhận (SĐT, địa chỉ…).']);
    addSheet('Ghi chu', notes);

    const filename = `800BEST_ThongKe_${safeName(state.office)}_${localStamp(now)}.xlsx`;
    XLSX.writeFile(wb, filename);
    log(`Đã xuất Excel: ${filename}`);
  }

  async function copyText(text) {
    try { await navigator.clipboard.writeText(text); return true; } catch { /* thử cách khác */ }
    const ta = document.createElement('textarea');
    ta.value = text; ta.style.cssText = 'position:fixed;opacity:0';
    document.body.appendChild(ta); ta.select();
    let ok = false;
    try { ok = document.execCommand('copy'); } catch { /* bỏ qua */ }
    ta.remove();
    return ok;
  }

  async function diagnose() {
    const c = getCalc();
    const stCol = state.orders ? colBy(state.orders.headers, /trạng thái/i) : -1;
    const info = {
      version: '3.11.3',
      signTimeKeys: state.sign ? state.sign.keys : null,
      signStats: state.sign ? { doneOrders: state.sign.okN, withSignTime: state.sign.signN } : null,
      weather: { on: state.wx.on, ok: !!state.wx.data, err: state.wx.err || '', hours: state.wx.data ? state.wx.data.times.length : 0 },
      page: location.pathname + location.hash,
      tabs: [...document.querySelectorAll('.el-tabs__item, [role="tab"], .ant-tabs-tab')].filter(visible).map(norm).slice(0, 20),
      employeeHeaders: state.emp ? state.emp.headers : null,
      employeeRows: state.emp ? state.emp.rows.length : 0,
      workCol: c && c.W ? c.W.workHeader : null,
      orderHeaders: state.orders ? state.orders.headers : null,
      orderRows: state.orders ? state.orders.rows.length : 0,
      statuses: state.orders && stCol >= 0 ? statusCounts(state.orders.rows, stCol) : null,
      overview: state.ov || null,
      office: state.office, province: state.province,
      history: { ready: state.histReady, error: state.histErr, records: state.hist.size, offices: [...new Set([...state.hist.values()].map(r => r.office))], lastSaved: state.lastSaved, pending: state.pending ? state.pending.reason : null },
      officeSelector: (() => {
        const i = findOfficeInput();
        return i ? { found: true, value: i.value, shownText: currentOfficeText(i), wrapperClass: selectWrap(i)?.className, placeholder: i.placeholder, itemsBelow: dropdownItems(i).slice(0, 8).map(norm), searchButtons: searchCandidates(i).slice(0, 5).map(x => ({ tag: x.el.tagName, cls: String(x.el.className).slice(0, 60), dy: Math.round(x.dy), dx: Math.round(x.dx) })), readOnly: i.readOnly, inputClass: i.className, wrapClass: i.parentElement?.className, wrapParentClass: i.parentElement?.parentElement?.className, searchButton: !!findSearchBtn(i) } : { found: false };
      })(),
      openDropdownOptions: getOptions().slice(0, 8).map(norm),
      endpoints: Object.fromEntries(state.endpoints)
    };
    const ok = await copyText(JSON.stringify(info, null, 2));
    log(ok ? 'Đã sao chép thông tin chẩn đoán (chỉ gồm tên cột, trạng thái, tên API; không có mã vận đơn hay thông tin người nhận).' : 'Không sao chép được vào clipboard.', !ok);
  }

  /* ------------------------------------------------------------------ *
   * Dựng giao diện
   * ------------------------------------------------------------------ */
  function mount() {
    if ($('#b8s-launch')) return;
    loadPrefs(); loadSites(); loadOffices(); loadKb().catch(() => {}); loadDisc();
    const css = `
#b8s-launch.logo{left:12px;top:12px;right:auto;bottom:auto;width:150px;padding:0;background:#fff;border-radius:8px;overflow:hidden;line-height:0;box-shadow:0 3px 12px #0005;border:1px solid #d6dbe4}#b8s-launch.logo .b8sl-t{display:block;line-height:1;padding:6px 0 0;text-align:center;color:#022344;font:800 12.5px/1 system-ui,Segoe UI,Arial,sans-serif;letter-spacing:3px}#b8s-launch.logo img{display:block;width:150px;height:86px}#b8s-launch.logo.bad{outline:3px solid #ef4444;outline-offset:-3px}#b8s-launch.logo:hover{filter:brightness(.96)}
#b8s-launch{position:fixed;left:50%;transform:translateX(-50%);right:auto;bottom:10px;z-index:999999;background:#4c3a8f;color:#fff;border:0;border-radius:10px;padding:12px 16px;font-weight:700;cursor:pointer;box-shadow:0 4px 18px #0005}
#b8s-disc{position:absolute;left:50%;transform:translateX(-50%);bottom:14px;z-index:30;width:min(920px,94%);max-height:60%;overflow:auto;box-shadow:0 8px 30px #0006;border-radius:10px;background:#fff}#b8s-disc .al{margin:0}#b8s-disc .b8sbtn.ghost{color:#374151;background:#f3f4f6;border:1px solid #d1d5db}
#b8s-app.lk .lk-hide{display:none}
#b8s-app{position:fixed;inset:2vh 2vw;z-index:1000000;display:none;flex-direction:column;border-radius:16px;overflow:hidden;box-shadow:0 20px 80px #0007;font:14px/1.45 "Segoe UI",Roboto,Arial,sans-serif;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:#1f2937}
#b8s-app *{box-sizing:border-box}
.b8s-top{padding:8px 18px 6px;color:#fff;display:flex;gap:12px;align-items:center;flex-wrap:wrap;justify-content:space-between}
.b8s-top h2{margin:0;font-size:22px;font-weight:800;letter-spacing:.3px;color:#ffe66d;text-shadow:0 2px 6px #2b1b5a99}
.b8s-sub{opacity:.85;font-size:12px;margin-top:2px}
.b8s-acts{display:flex;gap:8px;flex-wrap:wrap}
.b8s-pick{display:flex;gap:8px;flex:1;justify-content:center;min-width:280px}
.b8s-pick select{padding:8px 10px;border:0;border-radius:8px;min-width:220px;max-width:320px;flex:1;font-size:14px;font-weight:600;color:#4c3a8f;background:#fff}
.b8sbtn{border:0;border-radius:8px;background:#fff;color:#4c3a8f;padding:8px 13px;cursor:pointer;font-weight:700}
.b8sbtn.ghost{background:#ffffff33;color:#fff}
.b8s-tabs{display:flex;gap:6px;padding:0 20px}
.b8s-tab{border:0;background:#ffffff2e;color:#fff;padding:6px 16px;border-radius:9px 9px 0 0;cursor:pointer;font-weight:600}
.b8s-tab.on{background:#fff;color:#4c3a8f}
.b8s-main{flex:1;overflow:auto;padding:8px 16px;min-height:0}
.b8s-sec{color:#fff;font-weight:700;text-align:center;margin:8px 0;font-size:15px}
.b8s-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}
.b8s-card{background:#fffffff2;border-radius:12px;padding:10px 12px;text-align:center;box-shadow:0 2px 8px #0002}
#b8s-anm .kpibar{background:#fffffff2;border-radius:10px;padding:8px 16px;display:flex;flex-wrap:wrap;align-items:center;gap:8px 20px;box-shadow:0 2px 8px #0002;margin:0 0 8px}
#b8s-anm .kpibar .kl{flex:1;min-width:300px}#b8s-anm .kpibar .kh{display:flex;align-items:baseline;gap:10px;flex-wrap:wrap;font-size:12px;color:#64748b}
#b8s-anm .kpibar .kt{color:#4b5563;font-weight:700}#b8s-anm .kpibar .kn{font-size:24px;line-height:1.1}
#b8s-anm .kpibar .kr{display:flex;gap:26px;margin-left:auto;font-size:12px;color:#4b5563;text-align:right}#b8s-anm .kpibar .kr b{display:block;font-size:20px;color:#1f2937}
#b8s-anm .mk.kpi{background:#15803d;top:-5px;bottom:-5px;box-shadow:0 0 0 1px #fff}
.b8s-card span{display:block;font-size:12px;color:#4b5563}
.b8s-card b{display:block;font-size:26px;margin-top:2px;line-height:1.2}
.b8s-card small{display:block;font-size:11px;color:#6b7280;line-height:1.5}.b8s-card small i{font-style:normal;font-weight:700;color:#374151}
.c-red{color:#dc2626}.c-blue{color:#2563eb}.c-green{color:#16a34a}.c-amber{color:#d97706}.c-purple{color:#7c3aed}.c-teal{color:#0d9488}.c-orange{color:#ea580c}.c-pink{color:#db2777}
.b8s-tw{background:#fff;border-radius:12px;overflow:auto;margin-top:14px;max-height:60vh;box-shadow:0 2px 8px #0002}
.b8s-tw table,.b8s-panel table{width:100%;border-collapse:collapse;font-size:13px}
.b8s-tw th,.b8s-tw td,.b8s-panel th,.b8s-panel td{padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;white-space:nowrap}
.b8s-tw .l,.b8s-panel .l{text-align:left}
.b8s-tw thead th{position:sticky;color:#fff;cursor:pointer;font-weight:600;line-height:18px;z-index:1}
.b8s-tw thead tr:nth-child(1) th{top:0;height:34px;cursor:default}
.b8s-tw thead tr:nth-child(2) th{top:34px}
.g-nav{background:#1e3a8a}.g-gh{background:#3b82f6}.g-tt{background:#f59e0b}.g-sh{background:#4c3a8f}
.b8s-tw thead th.g-sh{color:#fff;font-weight:700;text-align:center;vertical-align:middle}
.sh-ld{display:inline-block;margin:0 0 0 6px;font:600 11px inherit;color:#4c3a8f;background:#fff;border:0;border-radius:99px;padding:2px 10px;cursor:pointer}
.sh-ld:disabled{cursor:default;opacity:.85}
.b8s-tw td small{display:block;font-size:10.5px;color:#6b7280;font-weight:400;line-height:1.2}
.b8s-tw td.sh-na{color:#9ca3af;font-size:11px}
/* Tổng quan gọn: mỗi dải chỉ số một hàng, bảng Shipper vừa đủ 15 dòng, từ ship 16 mới cuộn trong bảng */
.b8s-top{padding:6px 16px 4px}.b8s-top h2{font-size:19px}
.b8s-band{display:grid;grid-template-columns:96px repeat(7,minmax(0,1fr));gap:6px;margin-bottom:6px}
.b8s-band .bl{display:flex;align-items:center;justify-content:center;text-align:center;color:#fff;font-weight:700;font-size:12.5px;line-height:1.2;background:#ffffff26;border-radius:10px;padding:0 6px}
.b8s-band .b8s-card{padding:4px 6px 5px;border-radius:10px;min-width:0;box-shadow:0 1px 5px #0002}
.b8s-band .b8s-card span{font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.b8s-band .b8s-card b{font-size:21px;margin-top:0;line-height:1.15}
.b8s-band .b8s-card small{font-size:10.5px;line-height:1.2;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#b8s-dash .b8s-tw{margin-top:6px;max-height:540px}
#b8s-dash .b8s-tw th,#b8s-dash .b8s-tw td{padding:0 8px;height:30px}
#b8s-dash .b8s-tw thead tr:nth-child(1) th{height:28px;top:0}
#b8s-dash .b8s-tw thead tr:nth-child(2) th{height:30px;top:28px}
#b8s-dash .b8s-tw td small{display:inline;margin-left:5px;font-size:11px}
#b8s-dash .b8s-notes{margin-top:6px}
@media (max-width:900px){.b8s-band{grid-template-columns:repeat(2,minmax(0,1fr))}.b8s-band .bl{grid-column:1/-1;padding:4px}}

.b8s-tw tbody tr:nth-child(even){background:#f9fafb}
.b8s-tw tbody tr:hover{background:#eef2ff}
.rk1,.rk2,.rk3{color:#d97706;font-weight:700}
.r-ok{color:#16a34a;font-weight:700}.r-mid{color:#d97706;font-weight:700}.r-bad{color:#dc2626;font-weight:700}
.b8s-tw tr.total td{font-weight:700;background:#eef2f7;border-top:2px solid #c7d2e0;position:sticky;bottom:0}
.b8s-notes{margin-top:12px;color:#fff;font-size:12px;opacity:.95}
.b8s-empty{background:#fffffff2;border-radius:12px;padding:28px;text-align:center;color:#4b5563}
.b8s-panel{background:#fff;border-radius:12px;padding:14px;margin-bottom:12px}
.b8s-setrow{display:grid;grid-template-columns:minmax(240px,1fr) minmax(320px,2fr);gap:12px;align-items:start}@media(max-width:900px){.b8s-setrow{grid-template-columns:1fr}}
.b8s-panel h3{margin:0 0 10px;font-size:15px}
.b8s-panel label{display:block;font-size:12px;color:#4b5563}
.b8s-panel input[type=text],.b8s-panel input[type=number],.b8s-panel select{padding:6px;border:1px solid #cbd5e1;border-radius:6px;font-size:13px}
.row{display:flex;gap:14px;flex-wrap:wrap;align-items:flex-end;margin-bottom:10px}
.muted{color:#6b778c;padding:6px 0;font-size:12px}
#b8s-log{max-height:200px;overflow:auto;font:12px/1.5 monospace;background:#0f172a;color:#dbeafe;padding:8px;border-radius:7px}
.bad{color:#dc2626}#b8s-log .bad{color:#fecaca}
.b8s-warn{background:#fff7ed;color:#9a3412;border:1px solid #fdba74;border-radius:10px;padding:8px 12px;margin-bottom:10px;font-size:13px}
.b8sbtn.solid{background:#4c3a8f;color:#fff}
.b8s-tw small.up,.b8s-tw small.dn{display:block;font-size:10px;font-weight:600}
.b8s-tw small.up{color:#16a34a}.b8s-tw small.dn{color:#dc2626}
.b8sx{border:0;background:none;color:#9ca3af;cursor:pointer;font-size:14px}.b8sx:hover{color:#dc2626}
.b8s-panel input[type=date]{padding:5px;border:1px solid #cbd5e1;border-radius:6px;font-size:13px}
#b8s-hist .b8s-tw{margin-top:0}
#b8s-app .anm-top{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:8px}
#b8s-app .anm-f{color:#fff;font-size:12px;display:flex;gap:6px;align-items:center}
#b8s-app .anm-f input,#b8s-app .anm-f select{padding:6px;border:0;border-radius:8px;font-size:13px;font-weight:600;color:#4c3a8f}
#b8s-app .anm-meta{color:#fff;font-size:12px;margin-bottom:10px;opacity:.95}
#b8s-app .anm-wrap{position:relative}
#b8s-app .btn{display:inline-block;background:#4c3a8f;color:#fff;border-radius:8px;padding:7px 13px;font-size:12px;font-weight:700;cursor:pointer;white-space:nowrap;user-select:none}
#b8s-app .anm-top .btn{background:#fff;color:#4c3a8f}#b8s-app .btn.solid{background:#4c3a8f!important;color:#fff!important;padding:3px 10px;margin-left:8px}
#b8s-app .btn.ghost{background:#ffffff40!important;color:#fff!important}
#b8s-app .b8s-pop{position:absolute;right:0;top:34px;z-index:20;background:#fff;border:1px solid #c4b5fd;border-radius:10px;box-shadow:0 8px 24px #0004;padding:10px;width:330px;color:#1f2937}
#b8s-app .b8s-pop .pg{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:7px}#b8s-app .b8s-pop .pg b{width:100%;font-size:11px;color:#6b7280}
#b8s-app .b8s-pop .o{border:1px solid #cbd5e1;border-radius:12px;padding:3px 10px;font-size:12px;font-weight:600;cursor:pointer}#b8s-app .b8s-pop .o:hover{background:#ede9fe}#b8s-app .b8s-pop .o.on{background:#4c3a8f;color:#fff;border-color:#4c3a8f}
#b8s-app .b8s-pop .hn{font-size:11px;color:#6b7280;border-top:1px solid #e5e7eb;padding-top:6px;margin-top:4px}
#b8s-app .b8s-pop .ol{display:flex;align-items:center;gap:6px;padding:4px 2px;font-size:13px;cursor:pointer;color:#1f2937}
#b8s-app .b5cam{display:inline-block;margin-left:5px;padding:0 7px;border-radius:99px;background:#4c3a8f;color:#fff;font-size:10.5px;font-weight:700;cursor:pointer;white-space:nowrap}#b8s-app .b5cam:hover{background:#6d5bd0}
#b8s-bl{position:fixed;inset:0;z-index:60;background:#0008;display:flex;align-items:center;justify-content:center;padding:18px}#b8s-bl .blp{background:#fff;border-radius:12px;padding:14px;width:min(920px,100%);max-height:90vh;overflow:auto;box-shadow:0 10px 40px #0006;color:#1f2937}
#b8s-app .blp .blr{display:flex;align-items:center;gap:10px;padding:6px 4px;border-top:1px solid #e5e7eb;flex-wrap:wrap}#b8s-app .blp .blo{min-width:150px}#b8s-app .blp .bln{flex:1;min-width:200px;font-size:13px}#b8s-app .blp .blb{display:flex;gap:6px;flex-wrap:wrap}#b8s-app .blp .blb .btn{font-size:12px;padding:3px 10px}
#b8s-app .blp .blt{max-height:340px;overflow:auto;margin-top:8px;border:1px solid #e5e7eb;border-radius:8px}#b8s-app .blp .blt table{width:100%;border-collapse:collapse;font-size:12.5px}#b8s-app .blp .blt th{position:sticky;top:0;background:#4c3a8f;color:#fff;padding:5px 8px;text-align:left;white-space:nowrap}#b8s-app .blp .blt td{padding:4px 8px;border-top:1px solid #f1f5f9;white-space:nowrap}#b8s-app .blp .blt td.bla{white-space:normal;min-width:260px;max-width:420px}
#b8s-app .lnk{cursor:pointer}#b8s-app .lnk:hover{text-decoration:underline}
#b8s-app .dot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px}
#b8s-app .anm-two{display:grid;grid-template-columns:1.05fr 1fr;gap:12px;margin-top:12px}
@media(max-width:1000px){#b8s-app .anm-two{grid-template-columns:1fr}}
#b8s-app .anm-two .b8s-panel{margin-bottom:0}
#b8s-app .al-box{max-height:230px;overflow:auto}
#b8s-app .anm-hd{display:flex;justify-content:space-between;align-items:center;margin:12px 0 8px}
#b8s-app .anm-two h3{margin:0 0 8px}
#b8s-app .al{padding:6px 9px;border-radius:8px;margin-bottom:5px;font-size:12.5px}#b8s-app .al.r{background:#fee2e2;color:#991b1b}#b8s-app .al.a{background:#fef3c7;color:#92400e}#b8s-app .al.ok{background:#dcfce7;color:#166534}
#b8s-app .rk-row{display:grid;grid-template-columns:22px 150px 1fr 64px 56px;gap:8px;align-items:center;padding:5px 0;font-size:13px}
#b8s-app .rkn{display:inline-block;width:22px;height:22px;line-height:22px;text-align:center;border-radius:11px;background:#e5e7eb;font-weight:700;font-size:12px}
#b8s-app .rkb{height:14px;border-radius:7px;background:#e5e7eb;position:relative;overflow:hidden}#b8s-app .rkb i{position:absolute;left:0;top:0;bottom:0;border-radius:7px}
#b8s-app .rkd{font-size:11px;font-weight:700;text-align:left}
#b8s-app .up{color:#16a34a}#b8s-app .dn{color:#dc2626}#b8s-app .nt{color:#6b7280}
#b8s-app .tdl{display:block;font-size:10px;font-weight:700}
#b8s-app .b8s-card small.up{color:#16a34a}#b8s-app .b8s-card small.dn{color:#dc2626}
#b8s-app .chips{display:flex;gap:5px;flex-wrap:wrap;margin:6px 0}
#b8s-app .ch{border:1.5px solid var(--c,#4c3a8f);color:var(--c,#4c3a8f);border-radius:12px;padding:2px 9px;font-size:11px;font-weight:600;cursor:pointer;user-select:none}
#b8s-app .ch.off{opacity:.35;text-decoration:line-through}#b8s-app .ch.all{background:#4c3a8f;color:#fff;--c:#4c3a8f}
#b8s-app .b8s-tw td.h1{background:#dcfce7;color:#166534;font-weight:700}#b8s-app .b8s-tw td.h2{background:#fef3c7;color:#92400e;font-weight:700}#b8s-app .b8s-tw td.h3{background:#fee2e2;color:#991b1b;font-weight:700}
#b8s-anm .b8s-cards{gap:8px;grid-template-columns:repeat(6,1fr)}
#b8s-anm .anm-two{grid-template-columns:5fr 7fr}
#b8s-anm .anm-bot{display:grid;grid-template-columns:3fr 2fr;gap:8px;margin-top:8px;align-items:start}
#b8s-anm .anm-bot .b8s-tw{margin-top:0}
#b8s-anm .anm-alert{max-height:300px;overflow:auto}
@media(max-width:1000px){#b8s-anm .anm-bot,#b8s-anm .anm-two{grid-template-columns:1fr}}
#b8s-anm .b8s-card{padding:5px 8px;border-radius:10px}
#b8s-anm .b8s-card span{font-size:11px}
#b8s-anm .b8s-card b{font-size:22px;margin-top:0;line-height:1.15}
#b8s-anm .b8s-card small{font-size:10px;line-height:1.3}
#b8s-anm .anm-top{margin-bottom:4px;gap:8px}
#b8s-anm .anm-meta{margin-bottom:6px;font-size:11px}
#b8s-anm .anm-two{gap:8px;margin-top:8px}
#b8s-anm .b8s-panel{padding:8px 10px;border-radius:10px}
#b8s-anm .anm-two h3{font-size:13px;margin:0 0 4px}
#b8s-pick .pk{display:flex;flex-direction:column;gap:8px;font-size:13px}
#b8s-pick .pk-mu{color:#6b7280;font-size:11.5px}
#b8s-pick .pk-ctl{display:flex;flex-wrap:wrap;gap:8px;align-items:center;color:#fff}
#b8s-pick .pk-ctl label{display:inline-flex;gap:6px;align-items:center;font-weight:600}
#b8s-pick .pk-ctl input,#b8s-pick .pk-ctl select{font:600 13px inherit;border:0;border-radius:8px;padding:5px 9px;background:#fff;color:#4c3a8f}
#b8s-pick .pk-chip{background:#fff;color:#4c3a8f;font-weight:600;border-radius:8px;padding:6px 12px}
#b8s-pick .pk-chip small{font-weight:400;color:#6b7280}
#b8s-pick .pk-go{border:0;background:#fff;color:#4c3a8f;font-weight:600;border-radius:8px;padding:7px 14px;cursor:pointer}
#b8s-pick .pk-go:disabled{opacity:.6;cursor:default}
#b8s-pick .pk-st{color:#fff;font-size:12px}
#b8s-pick .pk-cards{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:8px}
#b8s-pick .pk-c{background:#fffffff2;border-radius:12px;padding:7px 10px 8px;text-align:center;box-shadow:0 2px 8px #0002;min-width:0}
#b8s-pick .pk-c>span{display:block;font-size:12px;color:#4b5563}
#b8s-pick .pk-c>b{display:block;font-size:26px;line-height:1.15;font-variant-numeric:tabular-nums;color:#2f55d4}
#b8s-pick .pk-c small{display:block;font-size:11px;color:#6b7280;line-height:1.4}
#b8s-pick .pk-c small .go{color:#4c3a8f}
#b8s-pick .pk-d{display:flex;justify-content:center;gap:6px;font-size:11px;margin-top:1px;white-space:nowrap}
#b8s-pick .pk-d em{font-style:normal;color:#6b7280}
#b8s-pick .pk-c.click{cursor:pointer;outline:2px solid transparent;transition:outline-color .15s}
#b8s-pick .pk-c.click:hover{outline-color:#fff}
#b8s-pick .pk-up{color:#16a34a;font-weight:700}#b8s-pick .pk-dn{color:#dc2626;font-weight:700}#b8s-pick .pk-fl{color:#6b7280}
#b8s-pick .pk-warn{color:#92400e;font-weight:700}
#b8s-pick .pk-als{display:flex;flex-wrap:wrap;gap:8px}
#b8s-pick .pk-al{border-radius:9px;padding:6px 12px;font-size:12.5px;flex:1 1 320px}
#b8s-pick .pk-al.bad{background:#fee2e2;color:#b91c1c}#b8s-pick .pk-al.warn{background:#fef3c7;color:#92400e}#b8s-pick .pk-al.info{background:#f3f0ff;color:#4c3a8f}
#b8s-pick .pk-box{background:#fffffff2;border-radius:12px;padding:9px 12px;box-shadow:0 2px 8px #0002;display:flex;flex-direction:column;gap:6px;min-width:0;color:#1f2937}
#b8s-pick .pk-box>header{display:flex;flex-wrap:wrap;gap:4px 10px;align-items:center;justify-content:space-between}
#b8s-pick .pk-box h2{margin:0;font-size:14.5px;font-weight:700}
#b8s-pick .pk-x{color:#4c3a8f;font-weight:600;text-decoration:none}
#b8s-pick .pk-sc{overflow:auto;max-height:420px}
#b8s-pick table{border-collapse:collapse;width:100%}
#b8s-pick th,#b8s-pick td{padding:5px 7px;border-bottom:1px solid #e5e7eb;vertical-align:middle;white-space:nowrap;text-align:left}
#b8s-pick th{position:sticky;top:0;z-index:1;font-size:11px;font-weight:600;color:#4b5563;background:#fafafa;white-space:normal;line-height:1.25;vertical-align:bottom}
#b8s-pick .n{text-align:right;font-variant-numeric:tabular-nums}
#b8s-pick td.nm{overflow:hidden;text-overflow:ellipsis;max-width:230px}
#b8s-pick tr[data-a]{cursor:pointer}
#b8s-pick tr[data-a]:hover td{background:#faf8ff}
#b8s-pick tr.sel td{background:#f3f0ff!important}
#b8s-pick tr.sel td:first-child{box-shadow:inset 3px 0 0 #4c3a8f}
#b8s-pick .pk-stack{display:flex;height:9px;border-radius:5px;overflow:hidden;background:#e5e7eb;min-width:90px}
#b8s-pick .pk-stack i{display:block;height:100%}
#b8s-pick .pk-pill{display:inline-block;font-weight:700;font-size:11.5px;padding:1px 7px;border-radius:99px;font-variant-numeric:tabular-nums}
#b8s-pick .pk-pill.good{background:#dcfce7;color:#15803d}#b8s-pick .pk-pill.warn{background:#fef3c7;color:#92400e}#b8s-pick .pk-pill.bad{background:#fee2e2;color:#b91c1c}#b8s-pick .pk-pill.na{background:#e5e7eb;color:#6b7280}
#b8s-pick .pk-tag{display:inline-block;font-size:10.5px;font-weight:700;padding:0 6px;border-radius:99px;margin-left:5px;line-height:1.5;background:#f3f0ff;color:#4c3a8f}
#b8s-pick .pk-tag.bd{background:#fef3c7;color:#92400e}
#b8s-pick .pk-low{display:grid;grid-template-columns:minmax(0,42fr) minmax(0,58fr);gap:8px;align-items:stretch}
#b8s-pick .pk-tiles{display:grid;grid-template-columns:1fr 1fr;gap:6px}
#b8s-pick .pk-tiles4{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px}
#b8s-pick .pk-m{border:1px solid #e5e7eb;border-radius:9px;padding:6px 10px;display:flex;flex-direction:column;justify-content:center;min-width:0;background:#fff}
#b8s-pick .pk-m.cur{background:#f3f0ff;border-color:#d9d0ff}
#b8s-pick .pk-m .k{font-size:11px;color:#4b5563;font-weight:600}
#b8s-pick .pk-m .v{font-size:21px;font-weight:700;line-height:1.2;font-variant-numeric:tabular-nums}
#b8s-pick .pk-m .s{font-size:11px;color:#6b7280}
#b8s-pick .pk-leg{display:flex;gap:12px;font-size:11px;color:#6b7280}
#b8s-pick .pk-leg span:before{content:"";display:inline-block;width:10px;height:3px;border-radius:2px;margin-right:5px;vertical-align:3px;background:var(--k)}
#b8s-pick .pk-leg span.pk-mu:before{display:none}
#b8s-pick .pk-foot{color:#fff;font-size:11.5px;opacity:.95}
#b8s-anm .anm-hd{margin:8px 0 4px}
#b8s-anm .al{padding:3px 8px;font-size:12.5px;margin-bottom:3px;line-height:1.35}
#b8s-anm .al-box{max-height:none}
#b8s-anm .rk-row{padding:2px 0;font-size:13px;grid-template-columns:20px 130px 1fr 58px 52px}
#b8s-anm .rkn{width:18px;height:18px;line-height:18px;font-size:11px}
#b8s-anm .rkb{height:11px}
#b8s-anm .chips{margin:3px 0}
#b8s-anm .ch{padding:1px 8px;font-size:11px}
#b8s-anm .muted{padding:2px 0;font-size:11px}
#b8s-anm .b8s-tw{margin-top:8px;max-height:none;overflow:visible}
#b8s-anm .b8s-tw table{font-size:13px}
#b8s-anm .b8s-tw th,#b8s-anm .b8s-tw td{padding:3px 6px}
#b8s-anm .b8s-tw thead th{position:static;line-height:15px}
#b8s-anm .b8s-tw tr.total td{position:static}
#b8s-anm .tdl{display:inline;font-size:9px;margin-left:4px}
#b8s-anm .b8s-notes{margin-top:6px;font-size:11px}
#b8s-anm .seg{display:inline-flex;background:#ede9fe;border-radius:8px;padding:2px}
#b8s-anm .seg span{padding:2px 9px;border-radius:6px;font-size:11px;font-weight:700;color:#5b21b6;cursor:pointer;user-select:none}
#b8s-anm .seg span.on{background:#4c3a8f;color:#fff}
#b8s-anm .rk-row.two{grid-template-columns:20px 120px 1fr 54px 108px 108px}
#b8s-anm .rk-row.two .rkb{overflow:visible}
#b8s-anm .rkh{display:grid;grid-template-columns:20px 120px 1fr 54px 108px 108px;gap:8px;font-size:10px;color:#64748b;font-weight:700;padding:0 0 2px}
#b8s-anm .mk{position:absolute;top:-3px;bottom:-3px;width:3px;border-radius:2px;background:#1f2937}
#b8s-anm .mk.w{background:#6366f1}
#b8s-anm .cd{font-size:11px;font-weight:700;white-space:nowrap}#b8s-anm .cd small{font-weight:500;color:#64748b;font-size:10px;margin-left:3px}
#b8s-anm .mv{font-size:10px;font-weight:700}
#b8s-anm .lg{display:inline-block;height:11px;border-left:3px solid;margin-right:4px;vertical-align:middle}
#b8s-anm .slots{display:flex;gap:5px;align-items:center;color:#fff;font-size:11px;margin:2px 0 6px;flex-wrap:wrap}
#b8s-anm .slot{padding:2px 9px;border-radius:12px;background:#ffffff33;font-weight:700;color:#fff;cursor:pointer;user-select:none}
#b8s-anm .slot.fut{opacity:.5;cursor:default}
#b8s-anm .slot.miss{background:#fecaca;color:#b91c1c;text-decoration:line-through;cursor:default}
#b8s-anm .slot.due{background:#fde68a;color:#92400e;cursor:default}
#b8s-anm .slot.part{background:#fef3c7;color:#92400e}
#b8s-anm .slot.live{background:#bbf7d0;color:#166534}
#b8s-anm .slot.on{background:#fff;color:#4c3a8f}
#b8s-anm .slot.live.on{background:#bbf7d0;color:#166534;outline:2px solid #fff}
#b8s-anm .b8s-card .cmp2{display:flex;justify-content:center;gap:10px;font-size:10px;color:#64748b;font-weight:600}
#b8s-anm .b8s-card .cmp2 small{display:inline;font-size:10px;font-weight:700}
#b8s-anm .rkp{margin-top:8px;overflow-x:auto}
#b8s-anm .rk2{display:grid;align-items:center;gap:8px;padding:2px 6px;border-radius:9px;white-space:nowrap;font-size:13px}
#b8s-anm .rk2:nth-child(even){background:#f8fafc}
#b8s-anm .rk2.hd{font-size:10.5px;color:#64748b;font-weight:700;background:none;padding-bottom:2px}
#b8s-anm .rk2 .rkb{height:11px;overflow:visible}
#b8s-anm .rk2 .rkt{font-weight:700;overflow:hidden;text-overflow:ellipsis}
#b8s-anm .med{display:inline-block;width:24px;height:24px;border-radius:12px;text-align:center;line-height:24px;font-weight:800;font-size:12px;color:#fff;background:#94a3b8}
#b8s-anm .med.m1{background:linear-gradient(135deg,#fbbf24,#d97706)}#b8s-anm .med.m2{background:linear-gradient(135deg,#cbd5e1,#64748b)}#b8s-anm .med.m3{background:linear-gradient(135deg,#fdba74,#b45309)}
#b8s-anm .pill{display:inline-block;border-radius:6px;padding:1px 8px;text-align:center}
#b8s-anm .pill.h1{background:#dcfce7;color:#166534}#b8s-anm .pill.h2{background:#fef3c7;color:#92400e}#b8s-anm .pill.h3{background:#fee2e2;color:#991b1b}
#b8s-anm .toplin{display:flex;justify-content:center;align-items:center;gap:10px;flex-wrap:wrap;padding:3px 8px 5px;font-size:13px;color:#1f2937}
#b8s-anm .toplin em{font-style:normal;color:#cbd5e1}#b8s-anm .toplin i{font-style:normal;font-weight:800}#b8s-anm .toplin b{color:#b45309}
#b8s-anm .toplin .lnk{font-weight:800;cursor:pointer}
#b8s-anm .anm-bot2 .anm-an{max-height:230px}
#b8s-anm.tiny .anm-meta,#b8s-anm.tiny .chips,#b8s-anm.tiny details.b8s-notes{display:none}
#b8s-anm.tiny .anm-bot2 .anm-an{max-height:140px}#b8s-anm.tiny .anm-bot2 svg{height:105px}
#b8s-anm.tiny .b8s-card{padding-top:4px;padding-bottom:4px}
#b8s-anm .b8s-notes>summary{cursor:pointer;font-weight:700;color:#fff;font-size:11px}
#b8s-anm details.b8s-notes{padding:4px 8px}
#b8s-app.anm .b8s-pick,#b8s-app.anm #b8s-export,#b8s-app.anm .b8s-sub{display:none}
#b8s-app.anm .b8s-top>div:first-child{flex:1;text-align:center;padding-left:96px}
#b8s-anm .rk2.hd{white-space:normal;line-height:1.2;align-items:end}
#b8s-anm .qs{margin-left:4px;font-size:10px;font-weight:500;color:#64748b}
#b8s-anm .fc b{font-size:14px;color:#111827}#b8s-anm .fc.base b{color:#6b7280;font-style:italic}#b8s-anm .fcs{margin-left:4px;font-size:10px;font-weight:500;color:#64748b}
#b8s-anm .sw{display:inline-block;width:38px;height:20px;border-radius:10px;background:#cbd5e1;position:relative;cursor:pointer;flex:none}
#b8s-anm .sw i{position:absolute;top:2px;left:2px;width:16px;height:16px;border-radius:50%;background:#fff;transition:left .15s}
#b8s-anm .sw.on{background:#4c3a8f}#b8s-anm .sw.on i{left:20px}#b8s-anm .sw.dis{opacity:.55}
#b8s-anm .anm-bot2{display:grid;grid-template-columns:2.5fr 1fr;gap:8px;margin-top:8px;align-items:stretch;min-height:var(--botmin,0)}
#b8s-anm .cod .cdh{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
#b8s-anm .cdchip{display:inline-block;border-radius:9px;padding:1px 10px;font-size:12px;font-weight:700;background:#f1f5f9;color:#64748b}
#b8s-anm .cdchip.on{background:#ede9fe;color:#4c3aa8}#b8s-anm .cdchip.rd{background:#fee2e2;color:#991b1b}
#b8s-anm .cdcols{display:grid;grid-template-columns:1.5fr 1fr;gap:14px;margin-top:6px;flex:1;min-height:0;align-content:start}
#b8s-anm .cdt{display:grid;grid-template-columns:minmax(96px,1.3fr) minmax(98px,1fr) minmax(98px,1fr) minmax(98px,1fr) 56px 46px;gap:6px;align-items:center;padding:var(--cdpad,5px) 6px;font-size:14.5px;white-space:nowrap}
#b8s-anm .cdzero{margin-top:6px;border:1px solid #fecaca;background:#fff7f7;border-radius:8px;padding:6px 8px}
#b8s-anm .cdzl{max-height:150px;overflow:auto}
#b8s-anm .zc{cursor:pointer;text-decoration:underline;padding:0 3px;border-radius:4px}
#b8s-anm .zc.on{background:#fecaca}
#b8s-anm .cdt.hd{font-size:11.5px;color:#64748b;font-weight:700}
#b8s-anm .cdt:not(.hd):nth-child(even){background:#f8fafc}
#b8s-anm .cdt.tot{border-top:2px solid #e2e8f0;background:#f1f5f9!important;font-weight:800}
#b8s-anm .cdt .nm{font-weight:700;overflow:hidden;text-overflow:ellipsis}
#b8s-anm .cdt .rd,#b8s-anm .cdr .rd{color:#dc2626}#b8s-anm .cdt .gy{color:#94a3b8}
#b8s-anm .cdb{height:8px;border-radius:5px;background:#e5e7eb;position:relative;display:block}#b8s-anm .cdb i{position:absolute;left:0;top:0;bottom:0;border-radius:5px}
#b8s-anm .cdal{max-height:var(--cdal,215px);overflow:auto}#b8s-anm .cod{display:flex;flex-direction:column}
#b8s-anm .cda{display:grid;grid-template-columns:11px minmax(0,1fr) auto;gap:7px;align-items:center;padding:3px 6px;font-size:13.5px}
#b8s-anm .cda:nth-child(even){background:#fff7f7}
#b8s-anm .cda small{display:block;color:#64748b;font-size:11.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#b8s-anm .cdn{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#b8s-anm .cdr{text-align:right;white-space:nowrap}
#b8s-anm .cdd{width:11px;height:11px;border-radius:50%;background:#dc2626}#b8s-anm .cdd.o{background:#f97316}#b8s-anm .cdd.y{background:#eab308}
#b8s-anm details.cdchk table{width:100%;border-collapse:collapse;font-size:11.5px;background:#fff}#b8s-anm details.cdchk th{color:#475569;background:#f1f5f9;padding:2px 6px;text-align:right;font-weight:700;font-size:10.5px}#b8s-anm details.cdchk td{padding:2px 6px;text-align:right;border-top:1px solid #eef2f7;color:#1f2937}#b8s-anm details.cdchk th.l,#b8s-anm details.cdchk td.l{text-align:left}
#b8s-anm details.cdchk{margin-top:2px;font-size:12px}#b8s-anm .cod{padding-bottom:6px}#b8s-anm .cod h3{font-size:16px}#b8s-anm .cod .cdh b{font-size:14.5px!important}#b8s-anm details.cdchk>summary{cursor:pointer;font-weight:700;color:#4c3a8f;font-size:11px}
#b8s-anm.tiny .cdal{max-height:84px}#b8s-anm.tiny details.cdchk{display:none}#b8s-anm.tiny .cdt{padding:0 5px;font-size:11px}#b8s-anm.tiny .cod{padding-top:4px;padding-bottom:4px}#b8s-anm.tiny .cdcols{margin-top:2px}#b8s-anm.tiny .cod small.muted{display:none}
@media(max-width:1000px){#b8s-anm .cdcols{grid-template-columns:1fr}}
@media(max-width:1000px){#b8s-anm .anm-bot2{grid-template-columns:1fr}}
#b8s-anm .anm-an{max-height:420px;overflow:auto}
#b8s-anm .rsn{padding:6px 9px;border-radius:10px;margin-bottom:5px;cursor:pointer;border-left:5px solid;font-size:12.5px}
#b8s-anm .rsn.u{background:#f0fdf4;border-color:#16a34a}#b8s-anm .rsn.d{background:#fef2f2;border-color:#dc2626}
#b8s-anm .rsn .go{float:right;font-size:11px;color:#4c3a8f;font-weight:700}
#b8s-anm .rc{display:inline-block;font-size:11.5px;background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:1px 8px;margin:3px 4px 0 0;white-space:normal}
#b8s-anm .rsf{font-size:12px;color:#475569;padding:3px 0}
#b8s-anm .tiles{display:flex;gap:8px}
#b8s-anm .tile{flex:1;background:#f8fafc;border-radius:10px;padding:6px;text-align:center}
#b8s-anm .tile.off{opacity:.55}#b8s-anm .tile .tl{font-size:11px;color:#64748b}#b8s-anm .tile b{display:block;font-size:22px;line-height:1.2}#b8s-anm .tile small{font-size:11px;color:#475569}
#b8s-status{padding:6px 20px;color:#fff;font-size:12px;min-height:26px;background:#0000002e}
#b8s-status.bad{color:#fecaca}
@media(max-width:800px){#b8s-app{inset:1vh 1vw}.b8s-card b{font-size:21px}}`;
    const style = document.createElement('style');
    style.id = 'b8s-style';
    style.textContent = css;
    document.documentElement.appendChild(style);

    const launch = document.createElement('button');
    launch.id = 'b8s-launch';
    if (LOGO) { launch.classList.add('logo'); launch.innerHTML = `<span class="b8sl-t">THỐNG KÊ</span><img alt="Thống kê" src="${LOGO}">`; launch.title = 'Thống kê 800BEST'; }
    else launch.textContent = '📊 Thống kê 800BEST';
    document.documentElement.appendChild(launch);

    const app = document.createElement('div');
    app.id = 'b8s-app';
    app.innerHTML = `
<div class="b8s-top">
  <div><h2 id="b8s-title">📊 Thống kê tổng quan</h2><div class="b8s-sub" id="b8s-sub">Chưa cập nhật</div></div>
  <div class="b8s-pick">
    <select id="b8s-office-in"></select>
    <button class="b8sbtn" data-act="switch" id="b8s-switch" title="Chuyển sang bưu cục đã chọn và tự tải toàn bộ số liệu. Chọn lại đúng bưu cục đang xem để tải lại.">Chuyển bưu cục</button>
  </div>
  <div class="b8s-acts">
    <button class="b8sbtn" data-act="export" id="b8s-export">⬇ Xuất Excel</button>
    <button class="b8sbtn ghost" id="b8s-close">Đóng ×</button>
  </div>
</div>
<div class="b8s-tabs">
  <button class="b8s-tab on" data-view="dash">Tổng quan</button>
  <button class="b8s-tab" data-view="anm">ANM</button>
  <button class="b8s-tab" data-view="pick">Pickup-sale</button>
  <button class="b8s-tab" data-view="set">Cài đặt</button>
</div>
<div id="b8s-disc" style="display:none"></div>
<div class="b8s-main">
  <section id="b8s-v-dash"><div id="b8s-dash"></div></section>
  <section id="b8s-v-hist" style="display:none"><div id="b8s-hist"></div></section>
  <section id="b8s-v-anm" style="display:none"><div id="b8s-anm"></div></section>
  <section id="b8s-v-pick" style="display:none"><div id="b8s-pick"></div></section>
  <section id="b8s-v-set" style="display:none">
    <div class="b8s-setrow">
    <div class="b8s-panel">
      <h3>KPI POD2400</h3>
      <p class="muted">Nhập KPI để tab ANM hiện thanh KPI POD2400 của khu vực (màu: đạt KPI là xanh, thấp hơn dưới ${KPI_GAP} điểm là vàng, từ ${KPI_GAP} điểm trở lên là đỏ). Để trống thì thanh KPI ẩn. Số nhập chỉ lưu trong trình duyệt của bạn.</p>
      <div class="row"><label>Tên khu vực <input type="text" id="b8s-kpi-name" style="width:160px" placeholder="ví dụ: Hà Nội"></label> <label>KPI khu vực (%) <input type="number" id="b8s-kpi-reg" step="any" style="width:70px"></label></div>
      <div class="row"><label>KPI từng bưu cục (mỗi dòng: tên bưu cục = KPI, tên đúng như trong danh sách Bưu cục)<br><textarea id="b8s-kpi-off" rows="5" style="width:360px" placeholder="FS ABC = 77&#10;DS XYZ = 75"></textarea></label></div>
    </div>
    <div class="b8s-panel">
      <h3>TikTok Shop</h3>
      <div class="row"><label>Mã đầu vận đơn của TikTok Shop (cách nhau dấu phẩy)<br><input type="text" id="b8s-tt" style="width:260px"></label></div>
    </div>
    <div class="b8s-panel">
      <h3>🧑‍✈️ Ship đi làm</h3>
      <p class="muted">Chỉ ship đủ điều kiện mới được tính vào “DM làm / Đơn mỗi DM” và hiện trong bảng shipper (ở cả tab Tổng quan và tab ANM). Đơn của ship bị ẩn vẫn nằm trong các số tổng.</p>
      <div class="row"><label><input type="checkbox" id="b8s-dm-on"> Bật lọc ship đi làm</label></div>
      <div class="row"><label>Chức vụ được tính (mỗi dòng một chức vụ, đúng như cột “Chức vụ”)<br><textarea id="b8s-dm-depts" rows="4" style="width:360px"></textarea></label></div>
      <div class="row"><label>Quét đơn đầu tiên muộn hơn giờ này thì không tính <input type="text" id="b8s-dm-cut" style="width:70px" placeholder="14:00"></label></div>
      <div class="row"><label>Cuối ngày (từ 23:30 và khi xem cả ngày): giao thành công dưới <input type="number" id="b8s-dm-min" min="0" style="width:60px"> đơn thì không tính</label></div>
    </div>
    <div class="b8s-panel">
      <h3>🏢 Bưu cục quản lý</h3>
      <div class="muted">Tool tự tìm bưu cục khi bạn bấm tải số liệu. Muốn thêm thủ công: gõ đúng tên bưu cục (đúng cả chữ hoa, chữ thường và dấu) rồi bấm “Thêm”. Tool kiểm tra tên và chạy thử lấy đơn; bưu cục không tồn tại hoặc không có đơn nào (rỗng) sẽ không được thêm. “Mã COD” là mã bưu cục trên trang quản lý COD (trường signSiteId của yêu cầu listCodPayment), không có thì bảng COD của bưu cục đó để trống.</div>
      <div class="row"><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" id="b8s-disc-on"> Tự tìm bưu cục đang quản lý (bưu cục có đơn hôm nay) khi bấm “Tải số liệu” ở ANM hoặc “Chuyển bưu cục” ở Tổng quan. Không bao giờ tự xóa bưu cục.</label></div>
      <div class="row"><input type="text" id="b8s-of-name" autocomplete="off" spellcheck="false" placeholder="Tên bưu cục, ví dụ: tên đúng như trong ô chọn Bưu cục của trang" style="width:260px"><button class="b8sbtn solid" id="b8s-of-add">Thêm</button></div>
      <div id="b8s-of-st" class="muted" style="min-height:18px"></div>
      <div id="b8s-of-list"></div>
    </div>
    </div>
    <div class="b8s-panel">
      <h3>Phân loại trạng thái đơn</h3>
      <p class="muted">Chọn mỗi trạng thái được tính là gì. “Chưa quét giao” là đơn đã nhập kho nhưng chưa quét phát; mọi đơn khác (đang giao, thất bại, hoàn…) là “Đã quét giao”. Lựa chọn được lưu cho lần sau. Các dòng ✔/⚠ dưới bảng ở trang Tổng quan cho biết số tự tính có khớp trang gốc không.</p>
      <div id="b8s-map" style="overflow:auto;max-height:40vh"></div>
    </div>
    <div class="b8s-panel lk-hide">
      <h3>Ngưỡng màu và cảnh báo (tab ANM)</h3>
      <div id="b8s-th" style="overflow:auto"></div>
    </div>
    <div class="b8s-panel lk-hide">
      <h3>Thời tiết (phân tích nguyên nhân POD2400)</h3>
      <div class="row"><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" id="b8s-wx-on"> Dùng lượng mưa theo giờ của Open-Meteo (miễn phí, không cần khóa)</label></div>
      <div class="row"><label>Vĩ độ<br><input type="number" step="any" id="b8s-wx-lat" style="width:120px"></label><label>Kinh độ<br><input type="number" step="any" id="b8s-wx-lon" style="width:120px"></label><button class="b8sbtn solid" id="b8s-wx-test">Tải thử</button></div>
      <div class="muted">Bản này tắt sẵn và để tọa độ Hà Nội; hãy nhập tọa độ nơi bạn làm việc (các bưu cục gần nhau dùng chung một điểm) rồi tích bật. Chỉ gửi tọa độ ra ngoài, không gửi số liệu bưu cục.</div>
      <div id="b8s-wx-st" class="muted"></div>
    </div>
    <div class="b8s-panel">
      <h3>Dữ liệu so sánh (hôm qua và tuần trước)</h3>
      <div class="muted">Tool tự lấy danh sách đơn của hôm qua và cùng thứ tuần trước từ trang Kanban để so với hôm nay đúng cùng giờ, và đối chiếu với số tổng của Kanban (lệch thì không dùng). Bản nén chỉ gồm giờ nhập kho, giờ ký nhận, cờ đơn mới và tên ship; không có mã vận đơn hay thông tin người nhận; tự xóa sau ${KB_KEEP} ngày.</div>
      <div class="row"><button class="b8sbtn solid" id="b8s-kb">Lấy lại ngay</button></div>
      <div id="b8s-kb-st" class="muted"></div>
    </div>
    <div class="b8s-panel">
      <h3>Sao lưu / khôi phục lịch sử</h3>
      <div class="muted">Tab Lịch sử đã bỏ, số liệu vẫn được lưu như cũ (bản ghi ngày và bản chụp theo mốc giờ). Sao lưu ra file để phòng khi xóa dữ liệu trình duyệt.</div>
      <div class="row"><button class="b8sbtn solid" id="b8s-bk">Sao lưu</button><button class="b8sbtn solid" id="b8s-rs">Khôi phục</button><input type="file" id="b8s-rs-file" accept=".json,application/json" style="display:none"><button class="b8sbtn" id="b8s-gohist" title="Bảng lịch sử theo kỳ, shipper tổng hợp, so sánh các bưu cục">Mở bảng lịch sử chi tiết (nâng cao)</button></div>
      <div id="b8s-bk-st" class="muted"></div>
    </div>
    <div class="b8s-panel lk-hide">
      <h3>Bản chụp theo mốc giờ (tab ANM)</h3>
      <div id="b8s-snapinfo" class="muted"></div>
    </div>
    <div class="b8s-panel">
      <h3>Công cụ <small id="b8s-ver" style="font-weight:400;opacity:.7"></small></h3>
      <div class="row"><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" id="b8s-fast"> Tải nhanh bằng API khi bấm “Tải số liệu” ở tab ANM (lỗi sẽ tự dùng cách thường)</label></div>
      <div class="row"><button class="b8sbtn" style="background:#4c3a8f;color:#fff" id="b8s-diag">Chẩn đoán API</button></div>
      <div id="b8s-log"></div>
    </div>
  </section>
</div>
<div id="b8s-status"></div>`;
    document.documentElement.appendChild(app);

    launch.onclick = () => { app.style.display = 'flex'; readOverview(); renderUI(); onEnterAnm(); };
    $('#b8s-close').onclick = () => { app.style.display = 'none'; };
    $('#b8s-export').onclick = exportExcel;
    $('#b8s-diag').onclick = diagnose;
    {
      const on = $('#b8s-wx-on'), lat = $('#b8s-wx-lat'), lon = $('#b8s-wx-lon');
      on.checked = state.wx.on; lat.value = state.wx.lat; lon.value = state.wx.lon;
      on.onchange = () => { state.wx.on = on.checked; savePrefs(); renderWxInfo(); if (on.checked) loadWeather(true); };
      const coord = () => {
        const a = parseFloat(lat.value), b = parseFloat(lon.value);
        if (!Number.isFinite(a) || Math.abs(a) > 90 || !Number.isFinite(b) || Math.abs(b) > 180) { lat.value = state.wx.lat; lon.value = state.wx.lon; log('Tọa độ không hợp lệ (vĩ độ −90..90, kinh độ −180..180).', true); return; }
        state.wx.lat = a; state.wx.lon = b; state.wx.data = null; savePrefs(); if (state.wx.on) loadWeather(true); renderWxInfo();
      };
      lat.onchange = coord; lon.onchange = coord;
      { const nm = $('#b8s-of-name'), ls = $('#b8s-of-list');
        $('#b8s-of-add').onclick = officesAdd; nm.addEventListener('keydown', e => { if (e.key === 'Enter') officesAdd(); });
        ls.addEventListener('click', e => { const b = e.target.closest && e.target.closest('[data-del]'); if (b && !b.disabled) officesRemove(b.dataset.del); });
        ls.addEventListener('change', e => { const i = e.target.closest && e.target.closest('input[data-cod]'); if (i) officesSaveCod(i.dataset.cod, i.value); }); }
      $('#b8s-wx-test').onclick = () => { if (!state.wx.on) { state.wx.on = true; on.checked = true; savePrefs(); } loadWeather(true); renderWxInfo(); };
      $('#b8s-bk').onclick = backupHistory;
      if ($('#b8s-kb')) $('#b8s-kb').onclick = () => { state.kb.tried = {}; kbEnsure(true); };
      $('#b8s-rs').onclick = () => $('#b8s-rs-file').click();
      $('#b8s-rs-file').onchange = e => { const f = e.target.files && e.target.files[0]; e.target.value = ''; if (f) restoreHistory(f); };
      $('#b8s-gohist').onclick = () => { state.view = 'hist'; renderUI(); };
    }
    { const v = $('#b8s-ver'); if (v) { const upd = () => { v.textContent = `v3.14.0 · tiêu đề API: ${Object.keys(state.apiHdr).length ? 'đã bắt từ trang' : 'dùng mặc định'}${window.__b8sn > 1 ? ` · ⚠ đang chạy ${window.__b8sn} bản script cùng lúc, hãy tắt bớt trong Tampermonkey` : ''}`; }; upd(); setInterval(upd, 1500); } }
    { const dc = $('#b8s-disc-on'); if (dc) { dc.checked = !!disc.on; dc.onchange = () => { disc.on = dc.checked; saveDisc(); }; } }
    { const fa = $('#b8s-fast'); if (fa) { fa.checked = !!state.fast; fa.onchange = () => { state.fast = fa.checked; savePrefs(); }; } }
    const doSwitch = async () => {
      if (state.busy || disc.busy) return;
      // Không còn tự dò bưu cục ở đây nữa: chỉ chuyển sang đúng bưu cục đã có sẵn trong danh sách (tải ở tab Tổng quan → "Tải danh sách bưu cục").
      const n = $('#b8s-office-in').value;
      if (n) switchOffice(n); else log(OFFICES.length ? 'Hãy chọn một bưu cục.' : 'Chưa có bưu cục nào được theo dõi. Bấm “Tải danh sách bưu cục” ở trên.', !!OFFICES.length);
    };
    $('#b8s-disc').addEventListener('click', e => {
      const b = e.target.closest('[data-d]'); if (!b) return;
      const a = b.dataset.d;
      if (a === 'undo') discUndo();
      else if (a === 'close') { disc.note = ''; disc.undo = null; discRender(); }
      else if (a === 'padd') discPickDone(true);
      else if (a === 'pskip') discPickDone(false);
      else if (a === 'pall' || a === 'pnone') { (disc.pick || []).forEach(c => { c.on = a === 'pall'; }); const h = $('#b8s-disc'); if (h) h.dataset.h = ''; discRender(); }
    });
    $('#b8s-disc').addEventListener('change', e => {
      const c = e.target.closest('input[data-dpick]'); if (!c || !disc.pick) return;
      const it = disc.pick[+c.dataset.dpick]; if (it) it.on = c.checked;
    });
    $('#b8s-switch').onclick = doSwitch;
    app.querySelectorAll('.b8s-tab').forEach(t => t.addEventListener('click', () => { state.view = t.dataset.view; renderUI(); onEnterAnm(); if (state.view === 'pick') onEnterPick(); }));

    $('#b8s-dash').addEventListener('click', e => {
      if (e.target.closest('[data-act="officeList"]')) { refreshOfficeList(); return; }
      if (e.target.closest('[data-act="force-save"]')) { forceSave(); return; }
      if (e.target.closest('button[data-shtog]')) { state.sh.hide = !state.sh.hide; renderDashboard(); return; }
      const shb = e.target.closest('button[data-sh]');
      if (shb) { shLoad(shb.dataset.sh).catch(er => { state.sh.msg = String(er && er.message || er); shRedraw(true); }); return; }
      const th = e.target.closest('th[data-k]');
      if (!th) return;
      const key = th.dataset.k;
      state.sort = state.sort.key === key ? { key, dir: -state.sort.dir } : { key, dir: key === 'name' ? 1 : -1 };
      renderDashboard();
    });

    $('#b8s-tt').value = state.tiktok;
    const kpiFill = () => { $('#b8s-kpi-name').value = KPI_REGION_NAME; $('#b8s-kpi-reg').value = Number.isFinite(KPI_REGION) ? KPI_REGION : ''; $('#b8s-kpi-off').value = Object.entries(KPI_OFFICE).map(([k, v]) => `${k} = ${v}`).join('\n'); };
    const kpiSave = () => {
      const n = parseFloat(String($('#b8s-kpi-reg').value).replace(',', '.'));
      KPI_REGION = Number.isFinite(n) ? n : NaN; KPI_REGION_NAME = $('#b8s-kpi-name').value.trim();
      const o = {};
      for (const l of $('#b8s-kpi-off').value.split('\n')) { const m = /^(.+?)\s*[=:\t]\s*([\d.,]+)\s*%?\s*$/.exec(l.trim()); if (m) { const v = parseFloat(m[2].replace(',', '.')); if (Number.isFinite(v)) o[m[1].trim()] = v; } }
      KPI_OFFICE = o;
      try { localStorage.setItem('b8s_kpi', JSON.stringify({ n: KPI_REGION_NAME, r: Number.isFinite(KPI_REGION) ? KPI_REGION : null, o: KPI_OFFICE })); } catch { /* bỏ qua */ }
      kpiFill(); renderUI();
    };
    kpiFill();
    for (const id of ['name', 'reg', 'off']) $('#b8s-kpi-' + id).addEventListener('change', kpiSave);
    const dmFill = () => { const r = dmRulesOf(state.dm); $('#b8s-dm-on').checked = r.on; $('#b8s-dm-depts').value = r.depts.join('\n'); $('#b8s-dm-cut').value = r.cutoff; $('#b8s-dm-min').value = r.minOk; };
    const dmSave = () => {
      state.dm = dmRulesOf({ on: $('#b8s-dm-on').checked, depts: $('#b8s-dm-depts').value.split('\n'), cutoff: $('#b8s-dm-cut').value, minOk: $('#b8s-dm-min').value });
      dmFill(); savePrefs(); state.calc = null; state.calcKey = ''; renderUI();
    };
    dmFill();
    for (const id of ['on', 'depts', 'cut', 'min']) $('#b8s-dm-' + id).addEventListener('change', dmSave);
    $('#b8s-tt').addEventListener('change', e => { state.tiktok = e.target.value.trim() || 'TTVN'; e.target.value = state.tiktok; savePrefs(); renderUI(); });
    $('#b8s-map').addEventListener('change', e => {
      const sel = e.target.closest('select[data-status]');
      if (!sel) return;
      state.statusMap[sel.dataset.status] = sel.value;
      savePrefs(); renderUI();
    });

    let rz = 0; window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => { if (state.view === 'anm') fitAnm(); }, 150); });
    $('#b8s-anm').addEventListener('click', ((...a) => onAnmClick(...a)));
    $('#b8s-anm').addEventListener('change', ((...a) => onAnmChange(...a)));
    $('#b8s-pick').addEventListener('click', onPickClick);
    $('#b8s-pick').addEventListener('change', onPickChange);
    $('#b8s-th').addEventListener('change', e => {
      const t = e.target, f = t.dataset && t.dataset.th;
      if (!f) return;
      const [k, key] = f.split('.'), v = parseFloat(t.value);
      if (!Number.isFinite(v)) { renderThresholds(); return; }
      state.th[k] = Object.assign({}, state.th[k], { [key]: v });
      savePrefs(); renderUI();
    });
    $('#b8s-th').addEventListener('click', e => { if (e.target.id === 'b8s-th-reset') { state.th = {}; savePrefs(); renderThresholds(); renderUI(); } });
    renderThresholds();
    $('#b8s-hist').addEventListener('change', onHistChange);
    $('#b8s-hist').addEventListener('click', onHistClick);
    renderOfficeList();
    renderUI();
    loadHistory();
    loadSnapLog();
    loadSnaps().then(() => { renderUI(); snapBadge(); });
    if (!state.snap.timer) { state.snap.timer = setInterval(((...a) => snapTick(...a)), 30000); setTimeout(((...a) => snapTick(...a)), 8000); }
    lkBoot();
  }

  // Khởi động an toàn: lỗi nào cũng hiện nút báo lỗi (thay vì im lặng), và tự dựng lại nếu trang web xóa mất nút
  function showFatal(e) {
    try {
      console.error('[800BEST]', e);
      if (document.getElementById('b8s-fatal') || !document.documentElement) return;
      const b = document.createElement('button');
      b.id = 'b8s-fatal';
      b.textContent = '⚠ Tool 800BEST lỗi: ' + String(e && e.message || e).slice(0, 120) + ' (bấm để sao chép)';
      b.style.cssText = 'position:fixed;right:20px;bottom:22px;z-index:2147483647;background:#b91c1c;color:#fff;border:0;border-radius:10px;padding:10px 14px;font:600 13px sans-serif;cursor:pointer;max-width:70vw';
      b.onclick = () => { try { navigator.clipboard.writeText(String(e && e.stack || e)); } catch { /* bỏ qua */ } };
      document.documentElement.appendChild(b);
    } catch { /* bỏ qua */ }
  }
  const safeMount = () => { try { mount(); } catch (e) { showFatal(e); } };
  try { hookNetwork(); } catch (e) { console.error('[800BEST] hookNetwork', e); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', safeMount, { once: true });
  else safeMount();
  setInterval(() => {
    if (document.body && !document.getElementById('b8s-launch') && !document.getElementById('b8s-fatal')) {
      document.getElementById('b8s-app')?.remove();
      document.getElementById('b8s-style')?.remove();
      safeMount();
    }
  }, 2500);
})();
