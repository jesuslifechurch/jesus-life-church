/* Jesus Life Church - shared site logic.
   Edit the SETTINGS block below. Everything else can stay as it is. */

/* ================= SETTINGS ================= */
const SITE = {
    // Paste your Google Apps Script web app URL here (the same one used in give.html).
    // Leave '' to show the built-in sample content only.
    sheetUrl: 'https://script.google.com/macros/s/AKfycbyS2sXmjsQhuRkOYTxznfNO9pA5USDj2jfZkBXl6DDfsqLtumFdGj6jdh4-aJp0YOcE/exec',

    // Church WhatsApp line for the floating chat button (digits only, with country code)
    whatsapp: '263775284134',

    venue: 'Hatcliffe Mabanana Blue Tent',
    city: 'Harare, Zimbabwe',
    mapsQuery: 'Hatcliffe Mabanana Blue Tent Harare Zimbabwe',

    // Weekly services (times are Harare time, 24-hour). dow: 0 = Sunday ... 6 = Saturday
    services: [
        { name: 'Sunday Service',  dow: 0, start: '10:30', end: '13:30', day: 'Sundays',    times: '10:30 AM – 1:30 PM' },
        { name: 'Midweek Service', dow: 3, start: '17:00', end: '19:30', day: 'Wednesdays', times: '5:00 PM – 7:30 PM' }
    ]
};

// Shown when the sheet is not connected yet, or cannot be reached.
const DEFAULT_CONTENT = {
    events: [
        { date: '2026-10-23', enddate: '', time: 'All night', title: 'Night of Visitation 2', description: 'An all-night service. Come expecting.', location: '' },
        { date: '2026-11-08', enddate: '', time: '', title: '5th Church Anniversary & Prophet’s Birthday Celebration', description: 'Join us as we celebrate five years of God’s faithfulness.', location: '' }
    ],
    leaders: [
        { group: 'Senior Leadership', name: 'Prophets M & B Chiyangwa', role: 'Senior Leaders', photo: '', bio: '' }
    ],
    sermons: [],
    devotionals: [],
    ministries: [],
    gallery: [],
    testimonies: [],
    products: [],
    settings: {}
};
/* ============================================ */

const JLC = (() => {
    const CAT = 2 * 3600e3; // Harare is UTC+2 all year
    const $ = (s, r) => (r || document).querySelector(s);
    const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));

    // Build a DOM element safely (text is never treated as HTML)
    function h(tag, props, ...kids) {
        const el = document.createElement(tag);
        Object.entries(props || {}).forEach(([k, v]) => {
            if (v == null || v === false) return;
            if (k === 'class') el.className = v;
            else if (k === 'text') el.textContent = v;
            else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
            else el.setAttribute(k, v);
        });
        kids.flat().forEach((c) => { if (c != null && c !== false) el.append(c.nodeType ? c : document.createTextNode(c)); });
        return el;
    }

    function safeUrl(u) {
        u = String(u || '').trim();
        return /^https?:\/\//i.test(u) ? u : '';
    }

    function photoUrl(p) {
        p = String(p || '').trim();
        if (!p) return '';
        if (/^https?:\/\//i.test(p)) return p;
        if (/^[\w\- .()]+\.(jpe?g|png|webp|gif)$/i.test(p)) return 'photos/' + encodeURIComponent(p);
        return '';
    }

    /* ---------- dates (Harare time) ---------- */
    const pad = (n) => String(n).padStart(2, '0');
    const isYmd = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
    function todayYmd() { return new Date(Date.now() + CAT).toISOString().slice(0, 10); }
    function utcOf(ymd) { const [y, m, d] = ymd.split('-').map(Number); return Date.UTC(y, m - 1, d); }
    function daysBetween(a, b) { return Math.round((utcOf(b) - utcOf(a)) / 864e5); }
    function addDays(ymd, n) { const d = new Date(utcOf(ymd) + n * 864e5); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
    function fmtDate(ymd, long) {
        return new Date(utcOf(ymd)).toLocaleDateString('en-GB', { weekday: long ? 'long' : 'short', day: 'numeric', month: long ? 'long' : 'short', year: 'numeric', timeZone: 'UTC' });
    }
    function rel(ymd) {
        const n = daysBetween(todayYmd(), ymd);
        if (n === 0) return 'Today';
        if (n === 1) return 'Tomorrow';
        if (n > 1) return 'In ' + n + ' days';
        return 'Past';
    }
    const evEnd = (ev) => (isYmd(ev.enddate) && ev.enddate >= ev.date) ? ev.enddate : ev.date;

    function upcoming(events) {
        const t = todayYmd();
        return (events || []).filter((e) => isYmd(e.date) && evEnd(e) >= t).sort((a, b) => a.date.localeCompare(b.date));
    }
    function past(events) {
        const t = todayYmd();
        return (events || []).filter((e) => isYmd(e.date) && evEnd(e) < t).sort((a, b) => b.date.localeCompare(a.date));
    }

    /* ---------- weekly services ---------- */
    function nextService() {
        const now = Date.now();
        const nowC = new Date(now + CAT);
        let best = null;
        SITE.services.forEach((s) => {
            const [sh, sm] = s.start.split(':').map(Number);
            const [eh, em] = s.end.split(':').map(Number);
            for (let off = 0; off < 8; off++) {
                const day = new Date(Date.UTC(nowC.getUTCFullYear(), nowC.getUTCMonth(), nowC.getUTCDate() + off));
                if (day.getUTCDay() !== s.dow) continue;
                const y = day.getUTCFullYear(), m = day.getUTCMonth(), d = day.getUTCDate();
                const start = Date.UTC(y, m, d, sh, sm) - CAT;
                const end = Date.UTC(y, m, d, eh, em) - CAT;
                if (end <= now) continue;
                if (!best || start < best.start) best = { service: s, start: start, end: end, live: now >= start };
                break;
            }
        });
        return best;
    }
    function fmtWhen(ms) {
        return new Date(ms).toLocaleString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', hour: '2-digit', minute: '2-digit', timeZone: 'Africa/Harare' });
    }

    function mapsUrl() { return 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(SITE.mapsQuery); }

    /* ---------- content from the Google Sheet ---------- */
    let _content = null;
    function loadContent() {
        if (_content) return _content;
        _content = (async () => {
            if (!SITE.sheetUrl) return DEFAULT_CONTENT;
            try {
                const c = sessionStorage.getItem('jlc_content');
                const t = Number(sessionStorage.getItem('jlc_content_t') || 0);
                if (c && Date.now() - t < 60000) return JSON.parse(c);
            } catch (e) { /* ignore */ }
            try {
                const ctl = new AbortController();
                const to = setTimeout(() => ctl.abort(), 10000);
                const res = await fetch(SITE.sheetUrl + '?action=content', { signal: ctl.signal });
                clearTimeout(to);
                const d = await res.json();
                if (!d || !d.ok) throw new Error('bad response');
                const out = {
                    events: d.events || [],
                    leaders: (d.leaders && d.leaders.length) ? d.leaders : DEFAULT_CONTENT.leaders,
                    sermons: d.sermons || [],
                    devotionals: d.devotionals || [],
                    ministries: d.ministries || [],
                    gallery: d.gallery || [],
                    testimonies: d.testimonies || [],
                    products: d.products || [],
                    settings: d.settings || {}
                };
                try { sessionStorage.setItem('jlc_content', JSON.stringify(out)); sessionStorage.setItem('jlc_content_t', String(Date.now())); } catch (e) { /* ignore */ }
                return out;
            } catch (e) {
                return DEFAULT_CONTENT;
            }
        })();
        return _content;
    }

    /* ---------- add to calendar (.ics) ---------- */
    function downloadIcs(ev) {
        const esc = (s) => String(s || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
        const start = ev.date.replace(/-/g, '');
        const end = addDays(evEnd(ev), 1).replace(/-/g, '');
        const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+Z$/, 'Z');
        const desc = [ev.time ? 'Time: ' + ev.time : '', ev.description || ''].filter(Boolean).join('\n');
        const lines = [
            'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Jesus Life Church//Website//EN', 'BEGIN:VEVENT',
            'UID:' + start + '-' + Math.random().toString(36).slice(2, 10) + '@jesuslifechurch',
            'DTSTAMP:' + stamp,
            'DTSTART;VALUE=DATE:' + start,
            'DTEND;VALUE=DATE:' + end,
            'SUMMARY:' + esc(ev.title),
            'DESCRIPTION:' + esc(desc),
            'LOCATION:' + esc(ev.location || SITE.venue + ', ' + SITE.city),
            'END:VEVENT', 'END:VCALENDAR'
        ];
        const blob = new Blob([lines.join('\r\n')], { type: 'text/calendar;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = (ev.title || 'event').replace(/[^\w]+/g, '-').slice(0, 40) + '.ics';
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
    }


    /* ---------- forms (saved into the Google Sheet) ---------- */
    async function postForm(action, record, honey) {
        if (!SITE.sheetUrl) return { ok: false, offline: true };
        try {
            const ctl = new AbortController();
            const to = setTimeout(() => ctl.abort(), 15000);
            const res = await fetch(SITE.sheetUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ action: action, website: honey || '', record: record }),
                signal: ctl.signal
            });
            clearTimeout(to);
            return await res.json();
        } catch (e) { return { ok: false, offline: true }; }
    }
    function waLink(text, number) {
        return 'https://wa.me/' + (number || SITE.whatsapp) + (text ? '?text=' + encodeURIComponent(text) : '');
    }

    /* ---------- floating WhatsApp button + announcement / live bar ---------- */
    function initChrome() {
        const css = document.createElement('style');
        css.textContent =
            '.wa-float{position:fixed;right:18px;bottom:18px;width:58px;height:58px;border-radius:50%;background:#25d366;display:grid;place-items:center;box-shadow:0 8px 24px rgba(0,0,0,.45);z-index:900;transition:transform .2s}' +
            '.wa-float:hover{transform:scale(1.08)}.wa-float svg{width:32px;height:32px;fill:#fff}' +
            '.notice{position:fixed;left:0;right:0;top:var(--hdr,70px);z-index:950;display:flex;align-items:center;justify-content:center;gap:14px;padding:9px 48px 9px 16px;background:linear-gradient(90deg,#4c1d95,#7c3aed);color:#fff;font-size:.92rem;font-weight:600;text-align:center}' +
            '.notice a{color:#fde68a;text-decoration:underline}' +
            '.notice.live{background:linear-gradient(90deg,#991b1b,#dc2626)}' +
            '.notice .dot{display:inline-block;width:10px;height:10px;border-radius:50%;background:#fff;margin-right:8px;animation:jlcpulse 1.4s infinite}' +
            '.notice button{position:absolute;right:12px;top:50%;transform:translateY(-50%);background:none;border:0;color:#fff;font-size:1.2rem;cursor:pointer;opacity:.8}' +
            '@keyframes jlcpulse{0%,100%{opacity:1}50%{opacity:.25}}' +
            '.nav-links .more{position:relative}.nav-links .more>a{cursor:pointer}' +
            '.nav-links .more ul{display:none;position:absolute;right:0;top:100%;min-width:230px;margin-top:10px;padding:8px;list-style:none;background:#13131d;border:1px solid rgba(255,255,255,.1);border-radius:14px;box-shadow:0 18px 40px rgba(0,0,0,.5)}' +
            '.nav-links .more:hover ul,.nav-links .more.open ul{display:block}' +
            '.nav-links .more ul a{display:block;padding:10px 14px;border-radius:10px;color:#cbd5e1;font-size:.92rem}' +
            '.nav-links .more ul a:hover{background:rgba(124,58,237,.2);color:#fff}' +
            '@media (max-width:860px){.nav-links .more ul{display:block;position:static;min-width:0;margin:0;padding:0 0 0 14px;background:none;border:0;box-shadow:none}.nav-links .more>a{pointer-events:none;color:#a78bfa!important;font-size:.8rem!important;letter-spacing:2px;text-transform:uppercase}.nav-links .more ul a{padding:12px 0;border-radius:0;border-bottom:1px solid rgba(255,255,255,.08)}}' +
            '.install{position:fixed;left:16px;bottom:18px;z-index:900;display:flex;align-items:center;gap:10px;max-width:calc(100% - 100px);padding:10px 14px;background:#1a1a28;border:1px solid rgba(124,58,237,.5);border-radius:16px;color:#fff;font-size:.88rem;box-shadow:0 10px 30px rgba(0,0,0,.5)}' +
            '.install button{border:0;border-radius:999px;padding:8px 14px;font-weight:700;cursor:pointer}' +
            '.install .go{background:#7c3aed;color:#fff}.install .no{background:none;color:#9ca3af;padding:6px}' +
            '@media (prefers-reduced-motion:reduce){.notice .dot{animation:none}.wa-float{transition:none}}';
        document.head.appendChild(css);

        const wa = h('a', { class: 'wa-float', href: waLink('Hello Jesus Life Church,'), target: '_blank', rel: 'noopener noreferrer', 'aria-label': 'Chat with us on WhatsApp', title: 'Chat with us on WhatsApp' });
        wa.innerHTML = '<svg viewBox="0 0 32 32" aria-hidden="true"><path d="M16.04 3C9.4 3 4 8.4 4 15.03c0 2.12.55 4.18 1.6 6L4 29l8.17-1.56a12.03 12.03 0 0 0 3.87.63h.01C22.68 28.07 28 22.67 28 16.04 28 9.4 22.68 3 16.04 3zm0 22.04h-.01c-1.2 0-2.37-.32-3.4-.93l-.24-.15-4.85.93.95-4.73-.16-.25a9.97 9.97 0 0 1-1.52-5.3c0-5.5 4.5-9.97 10.04-9.97 5.54 0 10.04 4.47 10.04 9.97s-4.5 10.43-10.85 10.43zm5.5-7.5c-.3-.15-1.78-.88-2.06-.98-.28-.1-.48-.15-.68.15-.2.3-.78.98-.95 1.18-.18.2-.35.22-.65.08-.3-.15-1.27-.47-2.42-1.5-.9-.8-1.5-1.78-1.67-2.08-.18-.3-.02-.46.13-.6.13-.14.3-.35.45-.52.15-.18.2-.3.3-.5.1-.2.05-.38-.03-.53-.07-.15-.68-1.63-.93-2.23-.25-.58-.5-.5-.68-.5h-.58c-.2 0-.53.08-.8.38-.28.3-1.05 1.03-1.05 2.5s1.08 2.9 1.23 3.1c.15.2 2.1 3.2 5.1 4.48.7.3 1.27.5 1.7.62.72.23 1.37.2 1.88.12.58-.08 1.78-.73 2.03-1.43.25-.7.25-1.3.18-1.43-.08-.13-.28-.2-.58-.35z"/></svg>';
        document.body.appendChild(wa);

        const hdr = document.querySelector('header');
        if (hdr) document.documentElement.style.setProperty('--hdr', hdr.offsetHeight + 'px');

        loadContent().then((c) => {
            const st = c.settings || {};
            let bar = null;
            if (st.live === 'yes' && safeUrl(st.livelink)) {
                bar = h('div', { class: 'notice live', role: 'status' }, h('span', null, h('i', { class: 'dot' }), 'LIVE NOW: '), h('a', { href: safeUrl(st.livelink), target: '_blank', rel: 'noopener noreferrer', text: 'Watch the service' }));
                bar.dataset.key = 'live';
            } else if (st.announcement) {
                const link = safeUrl(st.announcementlink);
                bar = h('div', { class: 'notice', role: 'status' }, h('span', { text: st.announcement }), link ? h('a', { href: link, target: '_blank', rel: 'noopener noreferrer', text: 'Learn more' }) : null);
                bar.dataset.key = st.announcement;
            }
            if (!bar) return;
            try { if (sessionStorage.getItem('jlc_notice_closed') === bar.dataset.key) return; } catch (e) { /* ignore */ }
            const x = h('button', { type: 'button', 'aria-label': 'Close', text: '×' });
            x.addEventListener('click', () => { bar.remove(); try { sessionStorage.setItem('jlc_notice_closed', bar.dataset.key); } catch (e) { /* ignore */ } });
            bar.append(x);
            document.body.appendChild(bar);
        });
    }


    /* ---------- "Install our app" prompt (Android) and tip (iPhone) ---------- */
    function initInstall() {
        const standalone = (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) || window.navigator.standalone === true;
        if (standalone) return;
        try { const t = Number(localStorage.getItem('jlc_install_hide') || 0); if (Date.now() - t < 30 * 864e5) return; } catch (e) { /* ignore */ }
        const hide = (box) => { box.remove(); try { localStorage.setItem('jlc_install_hide', String(Date.now())); } catch (e) { /* ignore */ } };

        let deferred = null;
        window.addEventListener('beforeinstallprompt', (e) => {
            e.preventDefault();
            deferred = e;
            const go = h('button', { class: 'go', type: 'button', text: 'Install' });
            const no = h('button', { class: 'no', type: 'button', 'aria-label': 'Not now', text: '✕' });
            const box = h('div', { class: 'install', role: 'dialog' }, h('span', { text: '📲 Add Jesus Life Church to your phone' }), go, no);
            go.addEventListener('click', async () => { box.remove(); deferred.prompt(); try { await deferred.userChoice; } catch (err) { /* ignore */ } deferred = null; });
            no.addEventListener('click', () => hide(box));
            document.body.appendChild(box);
        });

        // iPhone / iPad Safari has no install pop-up, so show a short how-to once
        const ua = navigator.userAgent || '';
        const iOS = /iPad|iPhone|iPod/.test(ua) && !window.MSStream;
        const safari = /Safari/.test(ua) && !/CriOS|FxiOS|EdgiOS|OPiOS/.test(ua);
        if (iOS && safari) {
            const no = h('button', { class: 'no', type: 'button', 'aria-label': 'Close', text: '✕' });
            const box = h('div', { class: 'install', role: 'note' }, h('span', { text: '📲 To install: tap Share, then “Add to Home Screen”' }), no);
            no.addEventListener('click', () => hide(box));
            setTimeout(() => document.body.appendChild(box), 4000);
        }
    }

    /* ---------- YouTube helper ---------- */
    function ytId(url) {
        const m = String(url || '').match(/(?:youtube\.com\/(?:watch\?(?:.*&)?v=|embed\/|shorts\/)|youtu\.be\/)([\w-]{11})/);
        return m ? m[1] : '';
    }

    /* ---------- shared page behaviour ---------- */
    function initMenu() {
        const btn = $('#menuBtn'), links = $('#navLinks');
        if (!btn || !links) return;
        btn.addEventListener('click', () => {
            const open = links.classList.toggle('open');
            btn.setAttribute('aria-expanded', open ? 'true' : 'false');
            btn.textContent = open ? '✕' : '☰';
        });
        const more = links.querySelector('.more > a');
        if (more) more.addEventListener('click', (e) => { e.preventDefault(); more.parentElement.classList.toggle('open'); });
        links.addEventListener('click', (e) => {
            if (e.target.tagName === 'A' && !e.target.matches('.more > a')) { links.classList.remove('open'); btn.textContent = '☰'; btn.setAttribute('aria-expanded', 'false'); }
        });
    }
    function initReveal() {
        const items = $$('.reveal');
        if ('IntersectionObserver' in window) {
            const io = new IntersectionObserver((entries) => {
                entries.forEach((en) => { if (en.isIntersecting) { en.target.classList.add('in'); io.unobserve(en.target); } });
            }, { threshold: 0.12 });
            items.forEach((el) => io.observe(el));
        } else { items.forEach((el) => el.classList.add('in')); }
    }
    document.documentElement.classList.add('js');
    document.addEventListener('DOMContentLoaded', () => {
        initMenu(); initReveal(); initChrome(); initInstall();
        if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
    });

    return { $, $$, h, safeUrl, photoUrl, pad, isYmd, todayYmd, utcOf, daysBetween, addDays, fmtDate, rel, evEnd, upcoming, past, nextService, fmtWhen, mapsUrl, loadContent, downloadIcs, ytId, postForm, waLink };
})();
