/* Accounts: email + one-time code sign-in and sign-up, approval by the admin, and a private cloud workspace per person.
   Uses Supabase. The key below is the public (publishable) key: it is safe in a web page because the database
   rules (row level security) only let each person read and write their own workspace. */
(function(){
  "use strict";
  var SB_URL = "https://ilzjcdrludwpbkzmiyzw.supabase.co";
  var SB_KEY = "sb_publishable_u7X_QMSta26EODUmdxxkFA_TwVCKaF3";
  var root = document.documentElement;
  var $ = function(id){ return document.getElementById(id); };
  var sb = window.supabase.createClient(SB_URL, SB_KEY, {auth:{persistSession:true, autoRefreshToken:true}});
  var mode = 'in', email = '', cooldown = 0, ws = null, me = null, profile = null;
  window.KR_AUTH = {sb:sb};

  function ss(k,v){ try{ if(v===undefined) return sessionStorage.getItem(k); if(v===null) sessionStorage.removeItem(k); else sessionStorage.setItem(k,v); }catch(e){} return null; }
  function msg(t, good){ var m=$('gMsg'); m.textContent = t||''; m.style.color = good ? '#6FCB8E' : '#FF8A7A'; }
  function show(which){
    ['gLoading','gForm1','gForm2','gPending'].forEach(function(id){ $(id).hidden = (id!==which); });
    $('gTabs').hidden = (which!=='gForm1');
  }
  function setMode(m){
    mode = m; $('gtIn').classList.toggle('on', m==='in'); $('gtUp').classList.toggle('on', m==='up');
    $('gNameWrap').hidden = (m!=='up'); $('gName').required = (m==='up');
    $('gSend').textContent = 'Email me a code';
    $('gLead').textContent = m==='up' ? 'Create your account. We email you a code to confirm it, then the admin approves your access.' : 'Enter your email and we will send you a sign-in code.';
    msg('');
  }
  $('gtIn').addEventListener('click', function(){ setMode('in'); });
  $('gtUp').addEventListener('click', function(){ setMode('up'); });

  function tick(){
    var b=$('gResend'); if(cooldown>0){ cooldown--; b.disabled=true; b.textContent='Resend code ('+cooldown+')'; setTimeout(tick,1000); }
    else { b.disabled=false; b.textContent='Resend code'; }
  }
  async function sendCode(){
    var opts = {shouldCreateUser: mode==='up'};
    if(mode==='up') opts.data = {full_name: $('gName').value.trim()};
    var r = await sb.auth.signInWithOtp({email:email, options:opts});
    return r.error;
  }
  $('gForm1').addEventListener('submit', async function(e){
    e.preventDefault(); msg('');
    email = $('gEmail').value.trim().toLowerCase();
    if(mode==='up' && !$('gName').value.trim()){ msg('Enter your name.'); return; }
    var b=$('gSend'); b.disabled=true;
    var err = await sendCode(); b.disabled=false;
    if(err){
      msg(/signups not allowed|not allowed for otp/i.test(err.message) ? 'No account found for this email. Choose Sign up.' : (/rate|seconds|limit/i.test(err.message) ? 'Too many codes requested. Wait a minute and try again.' : 'Could not send the code. '+err.message));
      return;
    }
    $('gSentTo').textContent = email; show('gForm2'); $('gCode').value=''; $('gCode').focus(); cooldown=30; tick();
  });
  $('gResend').addEventListener('click', async function(){
    msg(''); cooldown=30; tick(); var err = await sendCode(); if(err) msg('Could not resend. '+err.message); else msg('A new code is on its way.', true);
  });
  $('gBack').addEventListener('click', function(){ show('gForm1'); msg(''); });
  $('gForm2').addEventListener('submit', async function(e){
    e.preventDefault(); msg('');
    var b=$('gVerify'); b.disabled=true;
    var r = await sb.auth.verifyOtp({email:email, token:$('gCode').value.replace(/\s+/g,''), type:'email'});
    b.disabled=false;
    if(r.error || !r.data || !r.data.user){ msg('That code is wrong or has expired. Check it or request a new one.'); return; }
    show('gLoading'); boot(r.data.user);
  });
  async function signOut(){ ss('kr_ws',null); try{ await sb.auth.signOut(); }catch(e){} location.reload(); }
  document.addEventListener('click', function(e){
    var t = e.target; if(!t) return;
    if(t.id==='signOut' || t.id==='gPendingOut') signOut();
    if(t.id==='gRecheck'){ show('gLoading'); sb.auth.getUser().then(function(r){ if(r.data && r.data.user) boot(r.data.user); else show('gForm1'); }); }
    if(t.id==='wsBack'){ ss('kr_ws',null); location.reload(); }
  });

  async function getProfile(uid){
    for(var i=0;i<4;i++){
      var r = await sb.from('kr_profiles').select('*').eq('id',uid).maybeSingle();
      if(r.data) return r.data; await new Promise(function(ok){ setTimeout(ok,800); });
    }
    return null;
  }
  function state(t,bad){ var s=$('saveState'); if(!s) return; s.textContent=t; s.className='status-pill mono '+(bad?'pill-bad':'pill-dim'); }
  var timer=null, dirty=false, saving=false;
  window.KR_SAVE = function(){ dirty=true; state('Saving…'); clearTimeout(timer); timer=setTimeout(flush,700); };
  async function flush(){
    if(saving){ timer=setTimeout(flush,500); return; }
    if(!dirty) return; dirty=false; saving=true;
    var S = window.KR_STORE, r = await sb.from('kr_data').update({data:S, customer_count:(S.kr_roster||[]).length, updated_at:new Date().toISOString()}).eq('user_id',ws).select('user_id');
    saving=false;
    if(r.error || !r.data || !r.data.length){ dirty=true; state('Not saved, retrying…',true); timer=setTimeout(flush,4000); }
    else state(dirty ? 'Saving…' : 'Saved to your account');
  }
  document.addEventListener('visibilitychange', function(){ if(document.hidden && dirty){ clearTimeout(timer); flush(); } });
  window.addEventListener('beforeunload', function(e){ if(dirty || saving){ e.preventDefault(); e.returnValue=''; } });

  async function boot(user){
    me = user; profile = await getProfile(user.id);
    if(!profile){ show('gForm1'); msg('Your account is not ready yet. Try again in a moment.'); return; }
    if(profile.status!=='approved'){
      $('gPendTitle').textContent = profile.status==='blocked' ? 'Access turned off' : 'Waiting for approval';
      $('gPendText').textContent = profile.status==='blocked' ? 'The admin has turned off access for '+profile.email+'. Contact the admin if this is a mistake.' : 'Your email is confirmed. The admin needs to approve '+profile.email+' before you can use Kitchen Run.';
      $('gRecheck').hidden = (profile.status==='blocked');
      show('gPending'); return;
    }
    ws = (profile.is_admin && ss('kr_ws')) || user.id;
    var wsEmail = profile.email;
    if(ws!==user.id){ var p2 = await sb.from('kr_profiles').select('email,full_name').eq('id',ws).maybeSingle(); wsEmail = p2.data ? (p2.data.full_name ? p2.data.full_name+' ('+p2.data.email+')' : p2.data.email) : ws; }
    var d = await sb.from('kr_data').select('data').eq('user_id',ws).maybeSingle();
    if(d.error){ show('gForm1'); msg('Could not load your workspace. '+d.error.message); return; }
    window.KR_STORE = (d.data && d.data.data && typeof d.data.data==='object') ? d.data.data : {};
    window.KR_AUTH = {sb:sb, user:user, profile:profile, ws:ws};
    $('whoAmI').textContent = profile.full_name || profile.email;
    if(profile.is_admin){ var tb=document.querySelector('[data-tab="team"]'); if(tb) tb.hidden=false; }
    if(ws!==user.id){ $('wsBanner').hidden=false; $('wsWho').textContent = wsEmail; }
    state('Saved to your account');
    root.classList.add('authed');
    var s=document.createElement('script'); s.src='app.js';
    s.onload=function(){ if(profile.is_admin){ var t=document.createElement('script'); t.src='team.js'; document.body.appendChild(t); } };
    document.body.appendChild(s);
  }
  setMode('in'); show('gLoading');
  sb.auth.getSession().then(function(r){ if(r.data && r.data.session) boot(r.data.session.user); else show('gForm1'); });
})();
