// INSIGHT M2 DEMO -- NOT PRODUCTION CODE
// api/classlink.js — Vercel serverless function
// Proxies all ClassLink OneRoster API calls server-side so CORS is avoided.
// Called by the browser via fetch('/api/classlink?op=...')

const CLASSLINK_BASE    = 'https://classlinkcertification3-vn-v2.rosterserver.com';
const CLASSLINK_OR_BASE = CLASSLINK_BASE + '/ims/oneroster/v1p1';
const CLIENT_ID         = 'ea30f27a3e9ef691c1d32dcc';
const CLIENT_SECRET     = '1b184797c172ed9bee6fb481';

// NOTE: ClassLink docs say /oauth2/token — actual working endpoint is /token (HTTP Basic auth)
const TOKEN_URL = CLASSLINK_BASE + '/token';

// Entra email → ClassLink sourcedId (simulates DFC manual_map strategy)
const ENTRA_TO_CLASSLINK = {
    'teacher1@faronicsna.onmicrosoft.com': '5033_T5033-0005',
    'student1@faronicsna.onmicrosoft.com': '5033_S5033-0002',
    'student2@faronicsna.onmicrosoft.com': '5033_S5033-0003',
};

let cachedToken = null;

async function getToken() {
    if (cachedToken) return cachedToken;
    const creds = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64');
    const res = await fetch(TOKEN_URL, {
        method:  'POST',
        headers: { 'Authorization': `Basic ${creds}`, 'Content-Type': 'application/x-www-form-urlencoded' },
        body:    'grant_type=client_credentials'
    });
    if (!res.ok) throw new Error(`ClassLink token failed: HTTP ${res.status}`);
    const data = await res.json();
    cachedToken = data.access_token;
    setTimeout(() => { cachedToken = null; }, 55 * 60 * 1000); // clear before 1 h TTL
    return cachedToken;
}

async function clGet(path) {
    const token = await getToken();
    const res = await fetch(`${CLASSLINK_OR_BASE}${path}`, {
        headers: { 'Authorization': `Bearer ${token}` }
    });
    if (!res.ok) throw new Error(`ClassLink GET ${path} failed: HTTP ${res.status}`);
    return res.json();
}

module.exports = async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    if (req.method === 'OPTIONS') { res.status(200).end(); return; }

    const { op, email } = req.query;

    try {
        // ── op=teacher: get teacher's classes ─────────────────────────────────
        if (op === 'teacher') {
            const sourcedId = ENTRA_TO_CLASSLINK[(email || '').toLowerCase()];
            if (!sourcedId) return res.status(404).json({ error: 'No ClassLink mapping for: ' + email });

            const clsData  = await clGet(`/teachers/${encodeURIComponent(sourcedId)}/classes?status=active`);
            const classes  = clsData.classes || [];

            // Enrich with student counts in parallel
            const enriched = await Promise.all(classes.map(async cls => {
                let studentCount = 0;
                try {
                    const sd = await clGet(`/classes/${encodeURIComponent(cls.sourcedId)}/students`);
                    studentCount = (sd.users || []).length;
                } catch { /* non-fatal */ }
                return {
                    sourcedId:    cls.sourcedId,
                    title:        cls.title || cls.sourcedId,
                    courseCode:   cls.courseCode || null,
                    status:       cls.status || 'active',
                    studentCount
                };
            }));

            return res.status(200).json({ sourcedId, classes: enriched });
        }

        // ── op=student: get student's enrolled class ──────────────────────────
        if (op === 'student') {
            const sourcedId = ENTRA_TO_CLASSLINK[(email || '').toLowerCase()];
            if (!sourcedId) return res.status(404).json({ error: 'No ClassLink mapping for: ' + email });

            let enrolledClass = null;
            try {
                const cd = await clGet(`/students/${encodeURIComponent(sourcedId)}/classes`);
                const active = (cd.classes || []).filter(c => !c.status || c.status === 'active');
                if (active.length > 0) {
                    enrolledClass = { sourcedId: active[0].sourcedId, title: active[0].title || active[0].sourcedId };
                }
            } catch {
                // Fallback: scan teacher classes
                try {
                    const teacherSid = '5033_T5033-0005';
                    const tcd = await clGet(`/teachers/${encodeURIComponent(teacherSid)}/classes?status=active`);
                    for (const cls of (tcd.classes || [])) {
                        const sd = await clGet(`/classes/${encodeURIComponent(cls.sourcedId)}/students`);
                        if ((sd.users || []).some(s => s.sourcedId === sourcedId)) {
                            enrolledClass = { sourcedId: cls.sourcedId, title: cls.title || cls.sourcedId };
                            break;
                        }
                    }
                } catch { /* silent */ }
            }

            return res.status(200).json({ sourcedId, enrolledClass });
        }

        // ── op=sync: get counts for the sync summary panel ────────────────────
        if (op === 'sync') {
            const [td, cd, sd] = await Promise.all([
                clGet('/teachers?limit=200'),
                clGet('/classes?limit=200'),
                clGet('/students?limit=500')
            ]);
            return res.status(200).json({
                teacherCount: (td.users   || []).length,
                classCount:   (cd.classes || []).length,
                studentCount: (sd.users   || []).length,
                syncedAt:     new Date().toLocaleString()
            });
        }

        res.status(400).json({ error: 'Unknown op. Use op=teacher|student|sync' });
    } catch (err) {
        console.error('[classlink]', err.message);
        res.status(500).json({ error: err.message });
    }
};
