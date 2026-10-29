/* Cleanearth Site Audit – Supabase back end.
   Records live in one table ("records"), photos in a private storage bucket ("audit-photos").
   Exposes the small API the app uses: CE.db (collection/doc), CE.user, CE.assets, CE.downloads. */
(function(){
  'use strict';
  var C=window.CE_CONFIG||{};
  var TABLE='records',BUCKET='audit-photos';
  var sb=null,session=null,me=null;
  var cache={};        // "kind|audit|key" -> {data,mod}
  var people={};       // user id -> display name
  var scopes={};       // scope -> {loaded, channel}
  var subs=[];
  var currentAudit=null,pollTimer=null,polling=false;
  var photoURLs={},photoWaits={};

  function E(code,message){var e=new Error(message||code);e.code=code;return e}
  var ls={get:function(k,d){try{var v=localStorage.getItem(k);return v==null?d:JSON.parse(v)}catch(e){return d}},
          set:function(k,v){try{localStorage.setItem(k,JSON.stringify(v))}catch(e){}}};
  function wrap(res){
    if(res&&res.error){
      var er=res.error,msg=er.message||String(er);
      if(!navigator.onLine||/fetch|network|Failed to fetch|NetworkError|timeout/i.test(msg)||er.status>=500||er.status===429)throw E('unavailable',msg);
      if(er.code==='PGRST301'||er.status===401)throw E('unavailable','Your sign-in has expired. Reload the page.');
      throw E('invalid_argument',msg);
    }
    return res?res.data:null;
  }
  async function call(fn){try{return wrap(await fn())}catch(e){if(e.code)throw e;throw E('unavailable',(e&&e.message)||'No connection.')}}

  /* ---------- auth ---------- */
  function configured(){return !!(C.supabaseUrl&&C.supabaseAnonKey)}
  function nameOf(u){if(!u)return '';var m=u.user_metadata||{};return m.full_name||m.name||(u.email?u.email.split('@')[0].replace(/[._]/g,' ').replace(/\b\w/g,function(c){return c.toUpperCase()}):'')}
  async function init(){
    if(!configured())return {state:'unconfigured'};
    if(!window.supabase||!window.supabase.createClient)return {state:'error',message:'The sign-in library did not load. Check the connection and reload.'};
    sb=window.supabase.createClient(C.supabaseUrl,C.supabaseAnonKey,{auth:{persistSession:true,autoRefreshToken:true,detectSessionInUrl:false}});
    var r=await sb.auth.getSession();session=r&&r.data&&r.data.session;
    if(!session)return {state:'signedout'};
    me=session.user;people[me.id]=nameOf(me);
    sb.auth.onAuthStateChange(function(ev,s){if(ev==='SIGNED_OUT'){location.reload()}if(s)session=s});
    document.addEventListener('visibilitychange',function(){if(document.visibilityState==='visible')poll()});
    window.addEventListener('online',poll);
    return {state:'ready',name:nameOf(me),username:me.email||''};
  }
  async function signIn(email,password){
    await ensureClient();
    var r=await sb.auth.signInWithPassword({email:String(email).trim(),password:password});
    if(r.error){
      if(/invalid login|invalid credentials/i.test(r.error.message))throw E('invalid_argument','That email and password do not match. Check them, or ask the app manager to reset your password.');
      if(/fetch|network/i.test(r.error.message))throw E('unavailable','No connection. Sign in once you have signal; after that the app works offline.');
      throw E('invalid_argument',r.error.message);
    }
    location.reload();
  }
  async function ensureClient(){if(!sb&&configured()&&window.supabase)sb=window.supabase.createClient(C.supabaseUrl,C.supabaseAnonKey,{auth:{persistSession:true,autoRefreshToken:true,detectSessionInUrl:false}})}
  async function signOut(){try{await sb.auth.signOut()}catch(e){}try{Object.keys(localStorage).filter(function(k){return k.indexOf('ce-cache-')===0}).forEach(function(k){localStorage.removeItem(k)})}catch(e){}location.reload()}

  /* ---------- paths ---------- */
  function parse(path){
    var s=String(path).split('/');
    if(s[0]!=='audits'||s.length>4||s.some(function(x){return !x}))throw new TypeError('Unsupported path: '+path);
    if(s.length===1)return {coll:true,kind:'audits'};
    if(s.length===2)return {kind:'audits',audit:s[1],key:s[1]};
    if(s.length===3)return {coll:true,kind:s[2],audit:s[1]};
    return {kind:s[2],audit:s[1],key:s[3]};
  }
  function ck(p){return p.kind+'|'+p.audit+'|'+p.key}
  function ingest(row){
    var key=row.kind+'|'+row.audit_id+'|'+row.rec_key;
    if(row.updated_by&&row.updated_by_name)people[row.updated_by]=row.updated_by_name;
    cache[key]={data:row.data||{},mod:Date.parse(row.updated_at||0)||0};return key;
  }

  /* ---------- loading, realtime & subscriptions ---------- */
  function scopeFor(p){return p.kind==='audits'&&p.coll?'audits':'A:'+p.audit}
  function inScope(scope,key){var s=key.split('|');return scope==='audits'?s[0]==='audits':(s[1]===scope.slice(2)&&s[0]!=='audits')}
  function persist(scope){var o={};Object.keys(cache).forEach(function(k){if(inScope(scope,k))o[k]=cache[k]});ls.set('ce-cache-'+scope,o)}
  function restore(scope){var o=ls.get('ce-cache-'+scope,null);if(!o)return false;Object.keys(o).forEach(function(k){if(!cache[k])cache[k]=o[k]});return true}
  async function refresh(scope){
    var q=sb.from(TABLE).select('audit_id,kind,rec_key,data,updated_at,updated_by,updated_by_name');
    q=scope==='audits'?q.eq('kind','audits'):q.eq('audit_id',scope.slice(2)).neq('kind','audits');
    var rows=await call(function(){return q.limit(5000)});
    var seen={};(rows||[]).forEach(function(r){seen[ingest(r)]=1});
    Object.keys(cache).forEach(function(k){if(inScope(scope,k)&&!seen[k])delete cache[k]});
    scopes[scope].loaded=true;persist(scope);
  }
  function listen(scope){
    var ch=sb.channel('ce-'+scope.replace(/[^A-Za-z0-9_-]/g,'_'));
    var filter=scope==='audits'?'kind=eq.audits':'audit_id=eq.'+scope.slice(2);
    ch.on('postgres_changes',{event:'*',schema:'public',table:TABLE,filter:filter},function(pl){
      if(pl.eventType==='DELETE'){var o=pl.old||{};if(o.kind){delete cache[o.kind+'|'+o.audit_id+'|'+o.rec_key];persist(scope);notify()}else{poll()}return}
      if(pl.new){ingest(pl.new);persist(scope);notify()}
    }).subscribe();
    scopes[scope].channel=ch;
  }
  function snap(id,data){var d=data?Object.freeze(JSON.parse(JSON.stringify(data))):undefined;return {id:id,exists:!!data,data:function(){return d},metadata:{fromCache:false,hasPendingWrites:false}}}
  function docsFor(p){
    var out=[];Object.keys(cache).forEach(function(k){var s=k.split('|');if(s[0]===p.kind&&(p.kind==='audits'||s[1]===p.audit))out.push(snap(s[2],cache[k].data))});
    out.sort(function(a,b){return a.id<b.id?-1:a.id>b.id?1:0});return out;
  }
  function notify(force){
    subs.forEach(function(s){
      var docs=docsFor(s.parsed);var sig=JSON.stringify(docs.map(function(d){return [d.id,d.data()]}));
      if(!force&&sig===s.sig)return;s.sig=sig;
      try{s.cb({docs:docs,size:docs.length,empty:!docs.length,docChanges:function(){return []},metadata:{fromCache:false,hasPendingWrites:false}})}catch(e){console.error(e)}
    });
  }
  async function poll(){
    if(polling||!sb||!session)return;polling=true;
    try{for(var sc in scopes){try{await refresh(sc)}catch(e){if(e.code!=='unavailable')console.warn(e)}}notify()}
    finally{polling=false}
  }
  function collection(path){
    var p=parse(path);if(!p.coll)throw new TypeError('Not a collection: '+path);
    return {path:path,onSnapshot:function(cb,err){
      var sc=scopeFor(p);if(p.audit)currentAudit=p.audit;
      var s={path:path,parsed:p,cb:cb,err:err,sig:null};subs.push(s);
      if(!scopes[sc]){scopes[sc]={loaded:false};restore(sc);listen(sc)}
      setTimeout(function(){notify(true);refresh(sc).then(function(){notify()}).catch(function(e){if(e.code!=='unavailable'&&err)err(e)})},0);
      if(!pollTimer)pollTimer=setInterval(function(){if(document.visibilityState!=='hidden')poll()},60000);
      return function(){var i=subs.indexOf(s);if(i>=0)subs.splice(i,1)};
    }};
  }
  function rowFor(p,data){return {audit_id:p.audit,kind:p.kind,rec_key:p.key,data:data,updated_by:me&&me.id,updated_by_name:people[me&&me.id]||''}}
  function after(p,data){cache[ck(p)]={data:data,mod:Date.now()};persist(p.kind==='audits'?'audits':'A:'+p.audit);notify()}
  function doc(path){
    var p=parse(path);if(p.coll)throw new TypeError('Not a document: '+path);
    return {id:p.key,path:path,
      get:async function(){var r=await call(function(){return sb.from(TABLE).select('data').eq('audit_id',p.audit).eq('kind',p.kind).eq('rec_key',p.key).maybeSingle()});return snap(p.key,r?r.data:undefined)},
      set:async function(data){await call(function(){return sb.from(TABLE).upsert(rowFor(p,data),{onConflict:'audit_id,kind,rec_key'})});after(p,data)},
      // merge happens inside the database, so simultaneous answers from two people never overwrite each other
      merge:async function(patch){var d=await call(function(){return sb.rpc('merge_record',{p_audit:p.audit,p_kind:p.kind,p_key:p.key,p_patch:patch,p_name:people[me&&me.id]||''})});after(p,d||patch)},
      update:async function(patch){var d=await call(function(){return sb.rpc('merge_record',{p_audit:p.audit,p_kind:p.kind,p_key:p.key,p_patch:patch,p_name:people[me&&me.id]||''})});after(p,d||patch)},
      delete:async function(){await call(function(){return sb.from(TABLE).delete().eq('audit_id',p.audit).eq('kind',p.kind).eq('rec_key',p.key)});delete cache[ck(p)];persist('A:'+p.audit);notify()},
      onSnapshot:function(){return function(){}}};
  }

  // Save many records in one request (used for loading a whole checklist at once)
  async function setMany(items){
    if(!items.length)return;
    const rows=items.map(function(it){return rowFor(parse(it.path),it.data)});
    await call(function(){return sb.from(TABLE).upsert(rows,{onConflict:'audit_id,kind,rec_key'})});
    items.forEach(function(it){var p=parse(it.path);cache[ck(p)]={data:it.data,mod:Date.now()}});
    var sc={};items.forEach(function(it){var p=parse(it.path);sc[p.kind==='audits'?'audits':'A:'+p.audit]=1});Object.keys(sc).forEach(persist);notify();
  }
  // Delete a whole project: every record with its id, then its photos
  async function deleteAudit(id){
    await call(function(){return sb.from(TABLE).delete().eq('audit_id',id)});
    var removed=0;
    try{
      var folder=safe(id),all=[],offset=0;
      for(;;){var r=await sb.storage.from(BUCKET).list(folder,{limit:100,offset:offset});if(r.error||!r.data||!r.data.length)break;
        all=all.concat(r.data.map(function(f){return folder+'/'+f.name}));if(r.data.length<100)break;offset+=100}
      for(var i=0;i<all.length;i+=100){var rr=await sb.storage.from(BUCKET).remove(all.slice(i,i+100));if(!rr.error)removed+=Math.min(100,all.length-i)}
    }catch(e){}
    Object.keys(cache).forEach(function(k){if(k.split('|')[1]===id)delete cache[k]});
    try{localStorage.removeItem('ce-cache-A:'+id)}catch(e){}
    persist('audits');notify(true);
    return {photosRemoved:removed};
  }

  /* ---------- user, photos, files ---------- */
  var user={
    id:async function(){return me?me.id:null},
    can:async function(){return null},
    profiles:async function(ids){var o={};[].concat(ids).forEach(function(i){o[i]={id:i,name:people[i]||'',isMe:me&&i===me.id}});return o}
  };
  function safe(s){return String(s).replace(/[^A-Za-z0-9._-]/g,'_')}
  var assets={
    upload:async function(blob){
      var path=safe(currentAudit||'unsorted')+'/'+Date.now().toString(36)+'-'+Math.random().toString(36).slice(2,7)+'.jpg';
      await call(function(){return sb.storage.from(BUCKET).upload(path,blob,{contentType:blob.type||'image/jpeg',upsert:false})});
      photoURLs[path]=URL.createObjectURL(blob);
      return {id:path,url:photoURLs[path],sizeBytes:blob.size,contentType:blob.type||'image/jpeg'};
    }
  };
  function photoURL(id){
    if(photoURLs[id])return Promise.resolve(photoURLs[id]);
    if(!photoWaits[id])photoWaits[id]=call(function(){return sb.storage.from(BUCKET).download(id)})
      .then(function(b){photoURLs[id]=URL.createObjectURL(b);return photoURLs[id]}).catch(function(e){delete photoWaits[id];throw e});
    return photoWaits[id];
  }
  var downloads={save:async function(r){
    var blob=r.data instanceof Blob?r.data:new Blob([r.data]);var url=URL.createObjectURL(blob);
    var a=document.createElement('a');a.href=url;a.download=r.filename;document.body.appendChild(a);a.click();a.remove();
    setTimeout(function(){URL.revokeObjectURL(url)},30000);return {status:'saved'};
  }};

  window.CE={init:init,configured:configured,signIn:signIn,signOut:signOut,
    db:{collection:collection,doc:doc,setMany:setMany},deleteAudit:deleteAudit,user:user,assets:assets,downloads:downloads,photoURL:photoURL,refresh:poll};
})();
