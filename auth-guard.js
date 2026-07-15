// ── Supabase Auth Guard ──────────────────────────────────────────────
const SUPABASE_URL = 'https://jixaskajgkorrjhxiflc.supabase.co';
const SUPABASE_KEY = 'sb_publishable_5XGh3wyz-KPDh9Zl6liAsQ_yaB8F96X';

const _sb = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

// דפים ציבוריים שלא דורשים כניסה
const PUBLIC_PAGES = [
  'index.html', 'about.html', 'contact.html', 'services.html',
  'terms.html', 'articles.html', 'vba-guide.html', 'tools.html',
  'article-anomaly-visual.html','article-audit-reality-gap.html',
  'article-auditor-visit.html','article-contract-checklist.html',
  'article-excel-large-files.html','article-hours-overlap.html',
  'article-invoice-approval.html','article-money-no-check.html',
  'article-outlier-tool.html','article-procurement-audit.html',
  'article-professional-unit-trust.html','article-roi-controls.html',
  'article-sla-contract.html','article-tender-controls.html',
  'article-vendor-transparency.html',
  'login.html','register.html','reset-password.html','verify-email.html'
];

function _currentPage() {
  const p = window.location.pathname.split('/').pop();
  return p || 'index.html';
}

function _isPublicPage() {
  return PUBLIC_PAGES.includes(_currentPage());
}

function _isSubscriptionActive(profile) {
  if (!profile) return false;
  const status = profile.subscription_status;
  if (status === 'active') {
    if (!profile.subscription_end) return true;
    return new Date(profile.subscription_end) > new Date();
  }
  if (status === 'trial') {
    if (!profile.trial_end_date) return false;
    return new Date(profile.trial_end_date) > new Date();
  }
  return false;
}

async function _getProfile(userId) {
  const { data } = await _sb
    .from('profiles')
    .select('*')
    .eq('id', userId)
    .single();
  return data;
}

function _showBlockScreen(msg) {
  document.body.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:center;min-height:100vh;
      background:#f0f3f8;font-family:Heebo,sans-serif;direction:rtl">
      <div style="background:#fff;border-radius:12px;padding:40px;max-width:420px;
        text-align:center;box-shadow:0 4px 24px rgba(0,0,0,.1)">
        <div style="font-size:2.5rem;margin-bottom:16px">🔒</div>
        <h2 style="color:#0d1f3c;margin-bottom:12px">גישה חסומה</h2>
        <p style="color:#6b7a99;margin-bottom:24px">${msg}</p>
        <a href="contact.html" style="background:#1e6abf;color:#fff;padding:10px 24px;
          border-radius:8px;text-decoration:none;font-weight:700">צור קשר</a>
      </div>
    </div>`;
}

// ── פונקציה ראשית ────────────────────────────────────────────────────
async function _handleSession(session) {
  if (_isPublicPage()) return;

  if (!session) {
    const returnTo = encodeURIComponent(window.location.href);
    window.location.href = `login.html?return=${returnTo}`;
    return;
  }

  if (!session.user.email_confirmed_at) {
    window.location.href = 'verify-email.html';
    return;
  }

  const profile = await _getProfile(session.user.id);

  if (profile?.account_status === 'suspended') {
    _showBlockScreen('חשבונך הושהה. לפרטים פנה לתמיכה.');
    return;
  }

  // אם profile לא נטען — מניחים שהמנוי פעיל (שגיאת DB זמנית)
  if (!profile || _isSubscriptionActive(profile)) return;

  const returnTo = encodeURIComponent(window.location.href);
  window.location.href = `subscribe.html?return=${returnTo}`;
}

async function updateNav() {
  const { data: { session } } = await _sb.auth.getSession();
  const loginLink  = document.getElementById('nav-login');
  const logoutLink = document.getElementById('nav-logout');
  const dashLink   = document.getElementById('nav-dashboard');

  if (session) {
    if (loginLink)  loginLink.style.display  = 'none';
    if (logoutLink) logoutLink.style.display = '';
    if (dashLink)   dashLink.style.display   = '';
  } else {
    if (loginLink)  loginLink.style.display  = '';
    if (logoutLink) logoutLink.style.display = 'none';
    if (dashLink)   dashLink.style.display   = 'none';
  }
}

async function signOut() {
  await _sb.auth.signOut();
  window.location.href = 'index.html';
}

// ── הרצה דרך onAuthStateChange (פותר race condition) ─────────────────
let _guardDone = false;
let _guardTimer = null;

_sb.auth.onAuthStateChange((event, session) => {
  if (event === 'SIGNED_IN' || event === 'TOKEN_REFRESHED') {
    clearTimeout(_guardTimer);
    if (!_guardDone) { _guardDone = true; _handleSession(session); updateNav(); }
  } else if (event === 'INITIAL_SESSION') {
    if (session) {
      clearTimeout(_guardTimer);
      if (!_guardDone) { _guardDone = true; _handleSession(session); updateNav(); }
    } else {
      // אם דף ציבורי — לא צריך לחכות
      if (_isPublicPage()) { updateNav(); return; }
      // ממתינים ל-TOKEN_REFRESHED לפני החלטה
      _guardTimer = setTimeout(() => {
        if (!_guardDone) { _guardDone = true; _handleSession(null); updateNav(); }
      }, 2000);
    }
  } else if (event === 'SIGNED_OUT') {
    clearTimeout(_guardTimer);
    if (!_guardDone) { _guardDone = true; _handleSession(null); updateNav(); }
  }
});
