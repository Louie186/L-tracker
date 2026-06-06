// Lew's Tracker — Google Sheets Cloud Sync v8.5
var SHEET_API='https://script.google.com/macros/s/AKfycbye0SGVo8KhFEhtQwjiphVJye8oMNVGhrvK69c4NlB-TyvkgugKX8Gb31GyhtB6gT5fuw/exec';
var _syncTimer=null;
var _syncing=false;

function gsPost(payload){
  return fetch(SHEET_API,{
    method:'POST',
    redirect:'follow',
    headers:{'Content-Type':'text/plain;charset=utf-8'},
    body:JSON.stringify(payload)
  });
}

function gsGet(params){
  var url=SHEET_API+'?'+Object.keys(params).map(function(k){return k+'='+encodeURIComponent(params[k]);}).join('&');
  return fetch(url,{redirect:'follow'});
}

function gsPushAccounts(){
  if(!SHEET_API)return;
  gsPost({action:'saveAccounts',data:getUsers()}).then(function(r){
    return r.text();
  }).then(function(t){
    console.log('Accounts pushed:',t);
  }).catch(function(e){console.log('Account sync error:',e);});
}

function gsPullAccounts(){
  if(!SHEET_API)return;
  gsGet({action:'getAccounts'}).then(function(r){return r.json();}).then(function(cloud){
    if(cloud&&Array.isArray(cloud)&&cloud.length>0){
      var local=getUsers();
      var localNames=local.map(function(u){return u.username.toLowerCase();});
      var added=0;
      cloud.forEach(function(cu){
        if(cu&&cu.username&&localNames.indexOf(cu.username.toLowerCase())===-1){
          local.push(cu);
          added++;
        }
      });
      if(added>0){
        saveUsers(local);
        toastMsg(added+' account(s) synced from cloud');
      }
    }
  }).catch(function(e){console.log('Account pull error:',e);});
}

function gsPush(){
  if(!SHEET_API||!currentUser||_syncing)return;
  _syncing=true;
  gsPost({action:'saveData',user:currentUser,data:data}).then(function(r){
    return r.text();
  }).then(function(t){
    _syncing=false;
    console.log('Data pushed:',t);
    toastMsg('Synced to cloud');
  }).catch(function(e){_syncing=false;console.log('Sync error:',e);});
  gsPushAccounts();
}

function gsPull(){
  if(!SHEET_API||!currentUser)return;
  gsGet({action:'getData',user:currentUser}).then(function(r){return r.json();}).then(function(cloud){
    if(cloud&&cloud.tasks){
      data=Object.assign(structuredClone(data),cloud);
      save();renderAll();applySettings();
      toastMsg('Cloud data loaded');
    }
  }).catch(function(e){console.log('Pull error:',e);});
}

// Auto-sync on save (debounced 5 seconds)
var _origSave=window.save;
if(typeof _origSave==='function'){
  window.save=function(){
    _origSave();
    clearTimeout(_syncTimer);
    _syncTimer=setTimeout(gsPush,5000);
  };
}

// Sync accounts on register/remove
var _origSaveUsers=window.saveUsers;
if(typeof _origSaveUsers==='function'){
  window.saveUsers=function(u){
    _origSaveUsers(u);
    setTimeout(gsPushAccounts,500);
  };
}

// Init
setTimeout(function(){
  if(!SHEET_API)return;
  toastMsg('Cloud sync connected');
  gsPullAccounts();
  if(currentUser)gsPull();
},1500);
