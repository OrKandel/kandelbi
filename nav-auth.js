// nav-auth.js — updates the nav login link to show username after Supabase auth
(function(){
  const SUPABASE_URL='https://jixaskajgkorrjhxiflc.supabase.co';
  const SUPABASE_KEY='sb_publishable_5XGh3wyz-KPDh9Zl6liAsQ_yaB8F96X';

  function _client(){
    if(window._sb) return window._sb;
    if(window.supabase) return supabase.createClient(SUPABASE_URL,SUPABASE_KEY);
    return null;
  }

  function _apply(session){
    const link=document.querySelector('nav a[href="login.html"]');
    if(!link) return;
    const li=link.parentElement;

    if(session){
      const name=session.user.user_metadata?.full_name||session.user.email?.split('@')[0]||'משתמש';
      const first=name.split(' ')[0];
      link.textContent=first+' ▾';
      link.removeAttribute('href');
      link.style.cursor='pointer';
      li.style.position='relative';

      if(!li.querySelector('.nau-dd')){
        const dd=document.createElement('div');
        dd.className='nau-dd';
        dd.style.cssText='display:none;position:absolute;left:0;top:calc(100% + 4px);background:#fff;border:1px solid #dde3ed;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.12);min-width:175px;z-index:9999;padding:6px 0;font-family:Heebo,sans-serif;direction:rtl';
        dd.innerHTML=
          '<div style="padding:8px 14px;font-weight:700;color:#0d1f3c;border-bottom:1px solid #eee;font-size:.85rem;white-space:nowrap">'+name+'</div>'+
          '<a href="dashboard.html" style="display:block;padding:7px 14px;color:#1a2340;text-decoration:none;font-size:.85rem">⚙️ פרופיל / מנוי</a>'+
          '<a href="#" class="nau-logout" style="display:block;padding:7px 14px;color:#ef4444;text-decoration:none;font-size:.85rem">🚪 התנתקות</a>';
        li.appendChild(dd);

        link.addEventListener('click',function(e){
          e.stopPropagation();
          dd.style.display=dd.style.display==='none'?'block':'none';
        });
        document.addEventListener('click',function(){dd.style.display='none';});
        dd.querySelector('.nau-logout').addEventListener('click',async function(e){
          e.preventDefault();
          const c=_client();
          if(c) await c.auth.signOut();
          window.location.href='index.html';
        });
      }
    } else {
      // logged out — restore
      if(!link.getAttribute('href')) link.setAttribute('href','login.html');
      if(!link.textContent.includes('כניסה')) link.textContent='כניסה / הרשמה';
    }
  }

  function _init(){
    const c=_client();
    if(!c) return;
    c.auth.getSession().then(function(r){ _apply(r.data.session); }).catch(function(){});
    c.auth.onAuthStateChange(function(_,s){ _apply(s); });
  }

  if(document.readyState==='loading'){
    document.addEventListener('DOMContentLoaded',_init);
  } else {
    _init();
  }
})();
