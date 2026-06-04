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

// ── MSAL setup (v3) ────────────────────────────────────────────────────────────
let msalApp = null;

async function getMsalApp() {
    if (msalApp) return msalApp;
    // MSAL Browser v3: PublicClientApplication.createPublicClientApplication() is async
    msalApp = await msal.PublicClientApplication.createPublicClientApplication({
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
    if (typeof msal === 'undefined') {
        throw new Error('MSAL library failed to load. Please refresh the page and try again.');
    }
    const app    = await getMsalApp();
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

// ── Teacher sign-in ────────────────────────────────────────────────────────────
async function teacherSignIn() {
    const btn      = document.getElementById('btn-teacher-signin');
    const statusEl = document.getElementById('teacher-signin-status');
    btn.disabled   = true;
    document.getElementById('teacher-class-section').style.display = 'none';

    try {
        setStatus(statusEl, 'info', 'Opening Microsoft sign-in…');
        const msalResult = await signIn();

        setStatus(statusEl, 'info', `Signed in as ${esc(msalResult.email)} — fetching ClassLink classrooms…`);
        const clData = await clApi('teacher', msalResult.email);

        state.teacher = {
            email:        msalResult.email,
            displayName:  msalResult.displayName,
            sourcedId:    clData.sourcedId,
            classes:      clData.classes || [],
            selectedClass: clData.classes && clData.classes.length > 0 ? clData.classes[0] : null
        };

        setStatus(statusEl, 'success',
            `Signed in as <strong>${esc(msalResult.email)}</strong> (${esc(msalResult.displayName)})`);

        // Populate class dropdown
        const section = document.getElementById('teacher-class-section');
        const select  = document.getElementById('teacher-class-select');
        select.innerHTML = '';
        if (clData.classes && clData.classes.length > 0) {
            clData.classes.forEach((cls, i) => {
                const opt = document.createElement('option');
                opt.value = i;
                const cc  = cls.courseCode ? ` (${cls.courseCode})` : '';
                opt.textContent = `${cls.title}${cc} — ${cls.studentCount} student${cls.studentCount !== 1 ? 's' : ''}`;
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
document.addEventListener('DOMContentLoaded', () => showScreen(1));
