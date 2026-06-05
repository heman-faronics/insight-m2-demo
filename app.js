// INSIGHT M2 DEMO -- NOT PRODUCTION CODE
// app.js — browser-side demo logic
// Uses MSAL Browser (CDN) for Entra auth; calls /api/classlink for ClassLink data.

'use strict';

// ── Credentials ────────────────────────────────────────────────────────────────
const ENTRA_TENANT_ID = '04d60e3b-0fed-4071-a820-e4e87e8f8c98';
const ENTRA_CLIENT_ID = '7eb737a9-32f2-4298-b617-87d35051d95d';
const ENTRA_SCOPES    = ['openid', 'email', 'profile'];

// ── App state ──────────────────────────────────────────────────────────────────
const state = {
    currentScreen: 1,
    teacher: null,   // { email, displayName, sourcedId, classes, selectedClass }
    student: null,   // { email, displayName, sourcedId, enrolledClass }
    sync:    null    // { teacherCount, classCount, syncedAt }
};
const screenReady = { 1: false, 2: false, 3: true, 4: true, 5: true };

// ── MSAL setup (v2) ────────────────────────────────────────────────────────────
// Uses MSAL Browser v2 CDN (v3 dropped the UMD bundle so CDN use requires v2)
let msalApp = null;

function getMsalApp() {
    if (msalApp) return msalApp;
    if (typeof msal === 'undefined') {
        throw new Error('MSAL library failed to load. Check your network connection and refresh.');
    }
    msalApp = new msal.PublicClientApplication({
        auth: {
            clientId:    ENTRA_CLIENT_ID,
            authority:   `https://login.microsoftonline.com/${ENTRA_TENANT_ID}`,
            redirectUri: window.location.origin
        },
        cache: { cacheLocation: 'sessionStorage' }
    });
    return msalApp;
}

async function signIn() {
    const app    = getMsalApp();
    const result = await app.loginPopup({
        scopes: ENTRA_SCOPES,
        prompt: 'select_account'
    });
    const claims      = result.idTokenClaims || {};
    const email       = (claims.email || claims.preferred_username || result.account?.username || '').toLowerCase();
    const displayName = claims.name || email;
    return { email, displayName };
}

// ── ClassLink API helper (via Vercel serverless function) ──────────────────────
async function clApi(op, email) {
    const url = `/api/classlink?op=${op}${email ? '&email=' + encodeURIComponent(email) : ''}`;
    const res = await fetch(url);
    if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        throw new Error(body.error || `API error: HTTP ${res.status}`);
    }
    return res.json();
}

// ── Screen navigation ──────────────────────────────────────────────────────────
function navigate(delta) {
    const next = state.currentScreen + delta;
    if (next < 1 || next > 5) return;
    showScreen(next);
}

function showScreen(n) {
    document.querySelectorAll('.screen').forEach(el => el.classList.remove('active'));
    document.getElementById(`screen-${n}`).classList.add('active');
    state.currentScreen = n;

    for (let i = 1; i <= 5; i++) {
        const dot = document.getElementById(`dot-${i}`);
        const lbl = document.getElementById(`lbl-${i}`);
        const con = document.getElementById(`conn-${i}`);
        dot.classList.remove('active', 'done');
        lbl.classList.remove('active');
        if (con) con.classList.remove('done');
        if (i < n)  { dot.classList.add('done'); dot.innerHTML = '<i class="fas fa-check" style="font-size:.65rem"></i>'; if (con) con.classList.add('done'); }
        if (i === n) { dot.classList.add('active'); dot.textContent = i; lbl.classList.add('active'); }
        if (i > n)  { dot.textContent = i; }
    }

    if (n === 5) updateSummary();
    if (n === 3 && !state.sync) triggerSync();
    updateNav();
    window.scrollTo(0, 0);
}

function updateNav() {
    const n       = state.currentScreen;
    const backBtn = document.getElementById('btn-back');
    const nextBtn = document.getElementById('btn-next');
    const soBtn   = document.getElementById('btn-start-over');
    const hint    = document.getElementById('footer-hint');
    backBtn.style.display = n > 1 ? '' : 'none';
    soBtn.style.display   = n === 5 ? '' : 'none';
    if (n === 5) { nextBtn.style.display = 'none'; return; }
    nextBtn.style.display = '';
    nextBtn.disabled = !screenReady[n];
    hint.textContent = !screenReady[n]
        ? (n === 1 ? 'Complete teacher sign-in to continue' : n === 2 ? 'Complete student sign-in to continue' : '')
        : '';
}

// ── Tab switcher for real TC UI ────────────────────────────────────────────────
function tcSwitchTab(tab) {
    document.getElementById('tab-classids').classList.toggle('active', tab === 'classids');
    document.getElementById('tab-classrooms').classList.toggle('active', tab === 'classrooms');
    document.getElementById('tc-view-classids').style.display   = tab === 'classids'   ? '' : 'none';
    document.getElementById('tc-view-classrooms').style.display = tab === 'classrooms' ? '' : 'none';
}

// ── Teacher sign-in ────────────────────────────────────────────────────────────
async function teacherSignIn() {
    const btn      = document.getElementById('btn-teacher-signin');
    const statusEl = document.getElementById('teacher-signin-status');
    btn.disabled   = true;
    document.getElementById('teacher-class-section').style.display = 'none';

    try {
        setStatus(statusEl, 'info', 'Opening Microsoft sign-in…');
        const msalResult = await signIn();

        setStatus(statusEl, 'info', 'Fetching ClassLink classrooms…');
        const clData = await clApi('teacher', msalResult.email);

        state.teacher = {
            email:        msalResult.email,
            displayName:  msalResult.displayName,
            sourcedId:    clData.sourcedId,
            classes:      clData.classes || [],
            selectedClass: clData.classes && clData.classes.length > 0 ? clData.classes[0] : null
        };

        // Update the real product header: show user info row, hide plain logo
        document.getElementById('tc-logo-row').style.display = 'none';
        const userRow = document.getElementById('tc-user-row');
        userRow.style.display = 'flex';
        document.getElementById('tc-name-display').textContent  = msalResult.displayName;
        document.getElementById('tc-email-display').textContent = msalResult.email;

        // Enable START button, clear status
        statusEl.innerHTML = '';
        document.getElementById('btnStartSession').disabled = false;

        // Populate ClassLink dropdown
        const section = document.getElementById('teacher-class-section');
        const select  = document.getElementById('teacher-class-select');
        select.innerHTML = '';
        if (clData.classes && clData.classes.length > 0) {
            clData.classes.forEach((cls, i) => {
                const opt = document.createElement('option');
                opt.value = i;
                const cc  = cls.courseCode ? ` (${cls.courseCode})` : '';
                opt.textContent = `⟳ ${cls.title}${cc} — ${cls.studentCount} student${cls.studentCount !== 1 ? 's' : ''}`;
                select.appendChild(opt);
            });
            select.addEventListener('change', () => {
                state.teacher.selectedClass = clData.classes[+select.value];
                updateSummary();
            });
        } else {
            select.innerHTML = '<option>No active classes found in ClassLink</option>';
        }
        section.style.display = 'block';

        screenReady[1] = true;
        updateNav();
        updateSummary();
    } catch (err) {
        const msg = classifyError(err);
        setStatus(statusEl, 'danger', msg);
        btn.disabled = false;
    }
}

// ── Student sign-in ────────────────────────────────────────────────────────────
async function studentSignIn() {
    const btn      = document.getElementById('btn-student-signin');
    const statusEl = document.getElementById('student-signin-status');
    btn.disabled   = true;
    document.getElementById('student-class-section').style.display = 'none';

    try {
        setStatus(statusEl, 'info-light', 'Opening Microsoft sign-in…');
        const msalResult = await signIn();

        setStatus(statusEl, 'info-light', `Signed in as ${esc(msalResult.email)} — looking up enrolled class…`);
        const clData = await clApi('student', msalResult.email);

        state.student = {
            email:        msalResult.email,
            displayName:  msalResult.displayName,
            sourcedId:    clData.sourcedId,
            enrolledClass: clData.enrolledClass
        };

        setStatus(statusEl, 'success-light', `Signed in as <strong>${esc(msalResult.email)}</strong>`);

        const section   = document.getElementById('student-class-section');
        const classInfo = document.getElementById('student-class-info');
        const joinBtn   = document.getElementById('btn-student-join');

        if (clData.enrolledClass) {
            classInfo.innerHTML = `<div class="small text-white">You are enrolled in:<br><strong>${esc(clData.enrolledClass.title)}</strong></div>`;
            document.getElementById('student-class-id-field').value = 'cl_' + clData.enrolledClass.sourcedId;
            joinBtn.disabled = false;
        } else {
            classInfo.innerHTML = `<div class="small text-warning">No class assignment found.<br>Contact your teacher.</div>`;
            joinBtn.disabled = true;
        }
        section.style.display = 'block';

        screenReady[2] = true;
        updateNav();
        updateSummary();
    } catch (err) {
        const msg = classifyError(err);
        setStatus(statusEl, 'danger-light', msg);
        btn.disabled = false;
    }
}

// ── ClassLink sync (Screen 3) ──────────────────────────────────────────────────
async function triggerSync() {
    const btn     = document.getElementById('btn-sync-now');
    const spinner = document.getElementById('sync-spinner');
    const icon    = document.getElementById('sync-icon');
    const msgEl   = document.getElementById('sync-message');
    btn.disabled = true;
    icon.classList.add('d-none');
    spinner.classList.remove('d-none');
    msgEl.textContent = 'Connecting to ClassLink…';

    try {
        const data = await clApi('sync');
        state.sync = data;
        document.getElementById('last-sync-time').textContent      = data.syncedAt;
        document.getElementById('teacher-count').textContent        = data.teacherCount + ' teachers';
        document.getElementById('class-count').textContent          = data.classCount   + ' classes';
        document.getElementById('sync-status-badge').className      = 'badge bg-success';
        document.getElementById('sync-status-badge').innerHTML      = '<i class="fas fa-check me-1"></i>Success';
        msgEl.textContent = '';
    } catch (err) {
        document.getElementById('sync-status-badge').className  = 'badge bg-danger';
        document.getElementById('sync-status-badge').textContent = 'Failed';
        msgEl.textContent = err.message;
    } finally {
        btn.disabled = false;
        icon.classList.remove('d-none');
        spinner.classList.add('d-none');
    }
}

// ── Policy interactivity (Screens 4 & 5) ──────────────────────────────────────
function onTeacherAuthChange() {
    const mode  = document.querySelector('input[name="teacherAuthMode"]:checked')?.value;
    const isSso = mode === 'sso_preferred' || mode === 'sso_required';
    document.getElementById('sso-info-note').style.display       = isSso ? '' : 'none';
    document.getElementById('teacher-rostering-group').style.display = isSso ? '' : 'none';
}

function onStudentAuthChange() {
    const isSsoReq  = document.querySelector('input[name="studentAuthMode"]:checked')?.value === 'sso_required';
    const classField = document.getElementById('student-class-id-field');
    classField.disabled = isSsoReq;
    document.getElementById('class-id-hint').textContent = isSsoReq
        ? 'Class ID is assigned automatically from ClassLink roster.'
        : 'Enter the Class ID your teacher announced.';
    document.getElementById('auto-join-row').style.display = isSsoReq ? '' : 'none';
}

// ── Summary + Start Over ───────────────────────────────────────────────────────
function updateSummary() {
    if (state.teacher) {
        const cls = state.teacher.selectedClass;
        document.getElementById('summary-teacher-text').innerHTML =
            `Teacher: <strong>${esc(state.teacher.email)}</strong>` +
            (cls ? ` — class: <strong>${esc(cls.title)}</strong>` : '');
    }
    if (state.student) {
        const cls = state.student.enrolledClass;
        document.getElementById('summary-student-text').innerHTML =
            `Student: <strong>${esc(state.student.email)}</strong>` +
            (cls ? ` — enrolled in <strong>${esc(cls.title)}</strong>` : ' — no class found');
    }
}

function startOver() {
    state.teacher = null; state.student = null; state.sync = null;
    screenReady[1] = false; screenReady[2] = false;
    ['teacher-signin-status','student-signin-status'].forEach(id => document.getElementById(id).innerHTML = '');
    // Reset scenario switcher to scenario 1
    activeSimScenario = 1; activeSimStep = 1; simClassChosen = false; sim8ClassIdVisible = false;
    document.querySelectorAll('.scenario-pill').forEach((p, i) => p.classList.toggle('active', i === 0));
    const card = document.getElementById('scenario-card');
    if (card) card.textContent = SIM_SCENARIOS[0].card;
    renderSim(false);
    document.getElementById('btnStartSession').disabled = true;
    document.getElementById('teacher-class-section').style.display = 'none';
    document.getElementById('student-class-section').style.display = 'none';
    document.getElementById('btn-teacher-signin').disabled = false;
    document.getElementById('btn-student-signin').disabled = false;
    document.getElementById('last-sync-time').textContent     = '—';
    document.getElementById('teacher-count').textContent       = '—';
    document.getElementById('class-count').textContent         = '—';
    document.getElementById('sync-status-badge').className     = 'badge bg-secondary';
    document.getElementById('sync-status-badge').textContent   = 'Not synced';
    document.getElementById('student-class-id-field').value    = 'cl_5033_5033_01';
    document.getElementById('summary-teacher-text').innerHTML  = 'Teacher: <em>sign in on Screen 1 to populate</em>';
    document.getElementById('summary-student-text').innerHTML  = 'Student: <em>sign in on Screen 2 to populate</em>';
    showScreen(1);
}

// ── Helpers ────────────────────────────────────────────────────────────────────
function esc(s) {
    if (!s) return '';
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function setStatus(el, type, html) {
    const map = {
        'info':         'alert alert-info py-2 px-3 small mb-0',
        'success':      'alert alert-success py-2 px-3 small mb-0',
        'danger':       'alert alert-danger py-2 px-3 small mb-0',
        'info-light':   'small text-info',
        'success-light':'small text-success',
        'danger-light': 'small text-danger'
    };
    el.innerHTML = `<div class="${map[type] || 'small'}">${html}</div>`;
}

function classifyError(err) {
    const msg = (err.message || '').toLowerCase();
    if (msg.includes('popup') && msg.includes('close')) return 'Sign-in window was closed. Click to try again.';
    if (msg.includes('cancelled') || msg.includes('user_cancelled')) return 'Sign-in was cancelled.';
    if (msg.includes('no classlink mapping')) return err.message;
    if (msg.includes('network') || msg.includes('fetch')) return 'Network error. Check your connection.';
    return 'Sign-in failed: ' + esc(err.message || 'unknown error');
}

// ── Init ───────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    showScreen(1);
    initSimSwitcher();
});

// ══════════════════════════════════════════════════════════════════════════════
// SCENARIO SWITCHER — Screen 1 only
// activeSimScenario: 1–8   |   activeSimStep: 1–3
// ══════════════════════════════════════════════════════════════════════════════

let activeSimScenario = 1;
let activeSimStep     = 1;
let simClassChosen    = false;
let sim8ClassIdVisible = false;

const SIM_SCENARIOS = [
    {
        pill: '★ Class roster + sign-in',
        card: 'Your IT admin has set up Microsoft sign-in. Sign in and your classes load automatically — no need to type a Class ID.',
        steps: 3
    },
    {
        pill: 'Class roster required',
        card: "Microsoft sign-in is required and your class list must load from ClassLink to start. If ClassLink is unavailable, you can't start class.",
        steps: 3
    },
    {
        pill: 'Sign-in only',
        card: "Microsoft sign-in is required, but you'll still type your Class ID after signing in.",
        steps: 2
    },
    {
        pill: 'Sign-in or Class ID',
        card: "You can sign in with Microsoft to get your class list, or skip it and type your Class ID as usual.",
        steps: 3
    },
    {
        pill: 'Sign-in or Class ID (strict)',
        card: "Same as above, but if ClassLink is unavailable after sign-in, you won't be able to start class.",
        steps: 3
    },
    {
        pill: 'Sign-in or Class ID (no roster)',
        card: "You can sign in with Microsoft, but your class list won't load automatically — you'll still type your Class ID.",
        steps: 2
    },
    {
        pill: 'Class ID only',
        card: "No Microsoft sign-in. Enter your Class ID and optional password as usual. This is how Insight works today.",
        steps: 1
    },
    {
        pill: 'What if sign-in fails?',
        card: "See what the teacher experiences if Microsoft sign-in fails — and what recovery options are available.",
        steps: 3
    }
];

function initSimSwitcher() {
    const row = document.getElementById('scenario-pills-row');
    if (!row) return;
    row.innerHTML = SIM_SCENARIOS.map((s, i) =>
        `<button class="scenario-pill${i === 0 ? ' active' : ''}" onclick="switchSim(${i + 1})">${s.pill}</button>`
    ).join('');
    const card = document.getElementById('scenario-card');
    if (card) card.textContent = SIM_SCENARIOS[0].card;
    renderSim(false);
}

function switchSim(n) {
    activeSimScenario  = n;
    activeSimStep      = 1;
    simClassChosen     = false;
    sim8ClassIdVisible = false;

    // Pills
    document.querySelectorAll('.scenario-pill').forEach((p, i) =>
        p.classList.toggle('active', i + 1 === n));

    // Description card fade
    const card = document.getElementById('scenario-card');
    if (card) {
        card.classList.add('fading');
        setTimeout(() => {
            card.textContent = SIM_SCENARIOS[n - 1].card;
            card.classList.remove('fading');
        }, 100);
    }

    renderSim(true);
}

function simAdvanceStep() {
    simClassChosen = false;
    const maxSteps = SIM_SCENARIOS[activeSimScenario - 1].steps;
    if (activeSimStep < maxSteps) activeSimStep++;
    renderSim(true);
}

function simGoFallback() {
    simClassChosen = false;
    activeSimStep  = 3;
    renderSim(true);
}

function simRefresh() {
    const sel = document.getElementById('sim-class-sel');
    if (!sel) return;
    sel.style.opacity = '.35';
    setTimeout(() => { sel.style.opacity = ''; }, 700);
}

function simClassChange() {
    simClassChosen = true;
    const btn = document.getElementById('sim-start');
    if (btn) btn.disabled = false;
}

function simDynamicStart(inputId, btnId) {
    const inp = document.getElementById(inputId);
    const btn = document.getElementById(btnId);
    if (inp && btn) btn.disabled = !inp.value.trim();
}

function sim8RevealClassId() {
    sim8ClassIdVisible = true;
    renderSim(false);
}

function renderSim(animate) {
    const zone = document.getElementById('tc-sim-zone');
    if (!zone) return;
    const doRender = () => {
        zone.innerHTML = buildSimHtml();
        updateSimDots();
        zone.classList.remove('transitioning');
    };
    if (animate) {
        zone.classList.add('transitioning');
        setTimeout(doRender, 100);
    } else {
        doRender();
    }
}

function updateSimDots() {
    const dots = document.getElementById('tc-step-dots');
    if (!dots) return;
    const total = SIM_SCENARIOS[activeSimScenario - 1].steps;
    if (total <= 1) { dots.style.display = 'none'; return; }
    dots.style.display = 'flex';
    dots.innerHTML = Array.from({ length: total }, (_, i) =>
        `<span class="tc-dot${i + 1 === activeSimStep ? ' active' : ''}"></span>`
    ).join('');
}

// ── HTML fragment helpers ──────────────────────────────────────────────────────

const _MS_SVG = `<svg width="14" height="14" viewBox="0 0 21 21" style="flex-shrink:0"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>`;

function _msBtn(label, onclick) {
    return `<button class="sim-ms-btn" onclick="${onclick}">${_MS_SVG} ${esc(label)}</button>
            <div class="sim-simulated">(simulated — no real sign-in happens here)</div>`;
}

function _signedInPill() {
    return `<div class="sim-signed-pill"><span style="color:#16a34a">●</span> Signed in as teacher1@faronicsna.onmicrosoft.com</div>`;
}

function _classDropdown(showFallback) {
    const sel = `<select id="sim-class-sel" class="sim-select" onchange="simClassChange()">
      <option value="">— Choose a class —</option>
      <option>Science Lab · Period 2 · 14 students</option>
      <option>Mathematics 101 · Period 4 · 22 students</option>
      <option>English Literature · Period 6 · 18 students</option>
    </select>`;
    const bar = `<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
      <span class="sim-label" style="margin:0">Select your class:</span>
      <button class="sim-link" onclick="simRefresh()">⟳ Refresh list</button>
    </div>`;
    const fallback = showFallback
        ? `<div style="text-align:center;margin:6px 0">
             <button class="sim-link" style="color:#d97706" onclick="simGoFallback()">⚠ What if ClassLink is down?</button>
           </div>`
        : '';
    return `${bar}${sel}${fallback}
      <div style="text-align:right;margin-top:10px">
        <button class="sim-start-btn" id="sim-start" ${simClassChosen ? '' : 'disabled'}>Start Class</button>
      </div>`;
}

function _classIdEntry(label, inputId, btnId) {
    return `<label class="sim-label">${esc(label)}</label>
      <input id="${inputId}" type="text" class="sim-input" placeholder="e.g. Science101"
             oninput="simDynamicStart('${inputId}','${btnId}')">
      <div style="text-align:right">
        <button class="sim-start-btn" id="${btnId}" disabled>Start Class</button>
      </div>`;
}

function _orDivider() {
    return `<div class="sim-or">— or —</div>`;
}

function _msOrClassId(msOnclick, inputId, btnId) {
    return `${_msBtn('Sign in with Microsoft', msOnclick)}
      ${_orDivider()}
      ${_classIdEntry('Enter Class ID', inputId, btnId)}`;
}

// ── Main scenario HTML builder ─────────────────────────────────────────────────

function buildSimHtml() {
    const s    = activeSimScenario;
    const step = activeSimStep;

    // ── S1: SSO Required + Rostering + Allow Manual ─────────────────────────
    if (s === 1) {
        if (step === 1) return _msBtn('Sign in with Microsoft', 'simAdvanceStep()') +
            `<div class="sim-grey-text">Your school requires Microsoft sign-in</div>`;
        if (step === 2) return _signedInPill() + _classDropdown(true);
        if (step === 3) return _signedInPill() +
            `<div class="alert alert-warning py-2 px-3 small mb-3">⚠ Couldn't load your class list right now. Enter your Class ID to continue.</div>` +
            _classIdEntry('Enter your Class ID to continue', 'sim-cid-1', 'sim-sb-1');
    }

    // ── S2: SSO Required + Rostering + Block ────────────────────────────────
    if (s === 2) {
        if (step === 1) return _msBtn('Sign in with Microsoft', 'simAdvanceStep()') +
            `<div class="sim-grey-text">Your school requires Microsoft sign-in</div>`;
        if (step === 2) return _signedInPill() + _classDropdown(true);
        if (step === 3) return _signedInPill() +
            `<div class="alert alert-danger py-2 px-3 small mb-3">✕ Your class list couldn't be loaded and your school requires it to start class. Please contact your IT administrator.</div>
             <div style="text-align:right"><button class="sim-start-btn" disabled>Start Class</button></div>`;
    }

    // ── S3: SSO Required + No Rostering ─────────────────────────────────────
    if (s === 3) {
        if (step === 1) return _msBtn('Sign in with Microsoft', 'simAdvanceStep()') +
            `<div class="sim-grey-text">Your school requires Microsoft sign-in</div>`;
        if (step === 2) return _signedInPill() +
            _classIdEntry('Enter your Class ID to begin', 'sim-cid-3', 'sim-sb-3');
    }

    // ── S4: SSO Preferred + Rostering + Allow Manual ────────────────────────
    if (s === 4) {
        if (step === 1) return _msOrClassId('simAdvanceStep()', 'sim-cid-4a', 'sim-sb-4a');
        if (step === 2) return _signedInPill() + _classDropdown(true);
        if (step === 3) return _signedInPill() +
            `<div class="alert alert-warning py-2 px-3 small mb-3">⚠ Couldn't load your class list right now. Enter your Class ID to continue.</div>` +
            _classIdEntry('Enter your Class ID to continue', 'sim-cid-4b', 'sim-sb-4b');
    }

    // ── S5: SSO Preferred + Rostering + Block ───────────────────────────────
    if (s === 5) {
        if (step === 1) return _msOrClassId('simAdvanceStep()', 'sim-cid-5a', 'sim-sb-5a');
        if (step === 2) return _signedInPill() + _classDropdown(true);
        if (step === 3) return _signedInPill() +
            `<div class="alert alert-danger py-2 px-3 small mb-3">✕ Your class list couldn't be loaded and your school requires it to start class. Please contact your IT administrator.</div>
             <div style="text-align:right"><button class="sim-start-btn" disabled>Start Class</button></div>`;
    }

    // ── S6: SSO Preferred + No Rostering ────────────────────────────────────
    if (s === 6) {
        if (step === 1) return _msOrClassId('simAdvanceStep()', 'sim-cid-6a', 'sim-sb-6a');
        if (step === 2) return _signedInPill() +
            _classIdEntry('Enter your Class ID to begin', 'sim-cid-6b', 'sim-sb-6b');
    }

    // ── S7: Standard — Class ID only ────────────────────────────────────────
    if (s === 7) {
        return `<label class="sim-label">Enter your Class ID</label>
          <input id="sim-cid-7" type="text" class="sim-input" placeholder="e.g. Science101"
                 oninput="simDynamicStart('sim-cid-7','sim-sb-7')">
          <label class="sim-label">Password <span style="font-weight:400;color:#9ca3af">(optional)</span></label>
          <input type="password" class="sim-input" placeholder="Password">
          <div style="text-align:right">
            <button class="sim-start-btn" id="sim-sb-7" disabled>Start Class</button>
          </div>`;
    }

    // ── S8: Sign-in failure ──────────────────────────────────────────────────
    if (s === 8) {
        if (step === 1) return _msBtn('Sign in with Microsoft', 'simAdvanceStep()') +
            `<div class="sim-grey-text">Your school requires Microsoft sign-in</div>`;
        if (step === 2) return `
            <div class="alert alert-danger py-2 px-3 small mb-3">✕ Couldn't sign in with Microsoft. Check your internet connection and try again.</div>
            ${_msBtn('Try again with Microsoft', 'simAdvanceStep()')}
            <div style="text-align:center;margin-top:8px">
              <button class="sim-link" onclick="sim8RevealClassId()">Use Class ID instead →</button>
            </div>
            ${sim8ClassIdVisible ? `<div class="mt-3">${_classIdEntry('Enter your Class ID to continue', 'sim-cid-8', 'sim-sb-8')}</div>` : ''}`;
        if (step === 3) return _signedInPill() +
            `<div class="alert alert-success py-2 px-3 mb-3" style="font-size:11px">✓ Sign-in succeeded on retry.</div>` +
            _classDropdown(false);
    }

    return '';
}
