/* Admin Team tab: every person using the app, with approve, block, rename, open their workspace and reset. */
(function(){
  "use strict";
  var A = window.KR_AUTH, sb = A.sb, host = document.getElementById('teamHost'), pending = null;
  function esc(s){ return String(s==null?'':s).replace(/[&<>"]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
  function fmt(iso){ if(!iso) return ''; var d=new Date(iso); return isNaN(d)?'':d.toISOString().slice(0,16).replace('T',' '); }
  async function load(){
    host.innerHTML = '<div class="empty">Loading people…</div>';
    var a = await sb.from('kr_profiles').select('*').order('created_at',{ascending:false});
    var b = await sb.from('kr_data').select('user_id,customer_count,updated_at');
    if(a.error){ host.innerHTML = '<div class="empty">Could not load people. '+esc(a.error.message)+'</div>'; return; }
    var info = {}; (b.data||[]).forEach(function(r){ info[r.user_id]=r; });
    var rows = a.data, np = rows.filter(function(r){ return r.status==='pending'; }).length;
    var badge = document.getElementById('teamCount'); if(badge) badge.textContent = np ? String(np) : String(rows.length);
    host.innerHTML = '<div class="tbl-wrap"><table><thead><tr><th>Name</th><th>Email</th><th>Status</th><th>Customers</th><th>Last saved</th><th>Joined</th><th>Actions</th></tr></thead><tbody>'+
      rows.map(function(r){
        var d = info[r.id]||{}, me = r.id===A.user.id;
        var pill = r.is_admin ? '<span class="status-pill pill-ok">Admin</span>' : r.status==='approved' ? '<span class="status-pill pill-ok">Approved</span>' : r.status==='pending' ? '<span class="status-pill pill-warn">Pending</span>' : '<span class="status-pill pill-bad">Blocked</span>';
        var act = '';
        if(!r.is_admin){
          if(r.status!=='approved') act += '<button class="primary" data-a="approve" data-id="'+r.id+'">Approve</button>';
          if(r.status==='approved') act += '<button data-a="block" data-id="'+r.id+'">Block</button>';
          if(r.status==='blocked') act += '<button data-a="approve" data-id="'+r.id+'">Unblock</button>';
        }
        if(r.status==='approved') act += '<button data-a="open" data-id="'+r.id+'">'+(me?'My workspace':'Open workspace')+'</button>';
        if(!me) act += '<button class="danger" data-a="reset" data-id="'+r.id+'">'+(pending===r.id?'Confirm: erase data':'Reset data')+'</button>';
        return '<tr><td><input type="text" value="'+esc(r.full_name)+'" data-name="'+r.id+'" aria-label="Name" style="min-width:140px;"></td><td>'+esc(r.email)+'</td><td>'+pill+'</td><td class="mono">'+(d.customer_count||0)+'</td><td class="mono">'+esc(fmt(d.updated_at))+'</td><td class="mono">'+esc(fmt(r.created_at))+'</td><td><div class="act">'+act+'</div></td></tr>';
      }).join('')+'</tbody></table></div>';
  }
  host.addEventListener('click', async function(e){
    var b = e.target.closest('button[data-a]'); if(!b) return;
    var id = b.dataset.id, a = b.dataset.a;
    if(a==='open'){ try{ sessionStorage.setItem('kr_ws', id===A.user.id ? '' : id); if(id===A.user.id) sessionStorage.removeItem('kr_ws'); }catch(x){} location.reload(); return; }
    if(a==='reset'){
      if(pending!==id){ pending=id; load(); return; }
      pending=null; await sb.from('kr_data').update({data:{}, customer_count:0, updated_at:new Date().toISOString()}).eq('user_id',id); load(); return;
    }
    b.disabled = true;
    await sb.from('kr_profiles').update({status: a==='approve' ? 'approved' : 'blocked'}).eq('id',id); load();
  });
  host.addEventListener('change', async function(e){
    var i = e.target.closest('input[data-name]'); if(!i) return;
    await sb.from('kr_profiles').update({full_name:i.value.trim()}).eq('id',i.dataset.name);
  });
  document.getElementById('teamRefresh').addEventListener('click', load);
  document.querySelector('[data-tab="team"]').addEventListener('click', load);
  load();
})();
