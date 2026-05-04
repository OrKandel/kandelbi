// license-check.js — Cloudflare Worker
// בודק תוקף קבצי אקסל לפי מזהה ייחודי ותאריך הורדה

export default {
  async fetch(request, env) {
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Content-Type': 'application/json'
    };

    if (request.method === 'OPTIONS') return new Response(null, { headers: cors });
    if (request.method !== 'POST') return new Response('Method not allowed', { status: 405 });

    try {
      const { fileId, downloadDate, email, toolName } = await request.json();

      if (!fileId || !downloadDate) {
        return new Response(JSON.stringify({ valid: false, reason: 'מזהה חסר' }), { headers: cors });
      }

      // בדוק תאריך תוקף — 3 חודשים מתאריך ההורדה
      const downloaded = new Date(downloadDate);
      const expiry = new Date(downloaded);
      expiry.setMonth(expiry.getMonth() + 3);
      const now = new Date();

      if (now > expiry) {
        return new Response(JSON.stringify({
          valid: false,
          reason: 'פג תוקף',
          expiredOn: expiry.toLocaleDateString('he-IL'),
          message: `תוקף הקובץ פג בתאריך ${expiry.toLocaleDateString('he-IL')}. אנא הורד גרסה חדשה מ-kandelbi.org`
        }), { headers: cors });
      }

      // בדוק אם הקובץ לא בורשימה השחורה (KV store)
      if (env.REVOKED_IDS) {
        const revoked = await env.REVOKED_IDS.get(fileId);
        if (revoked) {
          return new Response(JSON.stringify({
            valid: false,
            reason: 'קובץ בוטל',
            message: 'קובץ זה בוטל. אנא צור קשר עם kandelbi.org'
          }), { headers: cors });
        }
      }

      // תקף!
      const daysLeft = Math.ceil((expiry - now) / (1000 * 60 * 60 * 24));
      return new Response(JSON.stringify({
        valid: true,
        daysLeft,
        expiry: expiry.toLocaleDateString('he-IL'),
        message: daysLeft <= 14 ? `שים לב: הקובץ יפוג בעוד ${daysLeft} ימים` : null
      }), { headers: cors });

    } catch (err) {
      // אם אין אינטרנט או שגיאת שרת — נאפשר עבודה (fallback לתאריך מקומי)
      return new Response(JSON.stringify({ valid: true, offline: true }), { headers: cors });
    }
  }
};
