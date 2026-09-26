/* Bandeau « Vous agissez en tant que … » affiché sur toutes les pages quand un administrateur a pris le rôle d'une personne. */
(function(){
  var SB="https://olqclfsywasktwirswpb.supabase.co",KEY="sb_publishable_Qry6l1DGbG9IkLtGT7CF4w_OyP4rI8x";
  function sess(){try{return JSON.parse(localStorage.getItem("kep.session")||"null");}catch(e){return null;}}
  function rpc(name,body){var s=sess();if(!s)return Promise.reject(new Error("non connecté"));
    return fetch(SB+"/rest/v1/rpc/"+name,{method:"POST",headers:{apikey:KEY,Authorization:"Bearer "+s.access_token,"Content-Type":"application/json"},body:JSON.stringify(body||{})})
      .then(function(r){return r.text().then(function(t){var d=null;try{d=t?JSON.parse(t):null;}catch(e){}if(!r.ok)throw new Error((d&&(d.message||d.msg))||("Erreur "+r.status));return d;});});}
  function esc(t){return String(t==null?"":t).replace(/[<>&"]/g,function(c){return {"<":"&lt;",">":"&gt;","&":"&amp;",'"':"&quot;"}[c];});}
  function bandeau(d){if(document.getElementById("kepComme"))return;
    var b=document.createElement("div");b.id="kepComme";b.setAttribute("role","status");
    b.style.cssText="position:sticky;top:0;align-self:start;justify-self:stretch;width:100%;box-sizing:border-box;z-index:9999;display:flex;flex-wrap:wrap;gap:8px 14px;align-items:center;justify-content:center;padding:8px 14px;background:#D1683F;color:#fff;font:600 13px/1.3 system-ui,-apple-system,'Segoe UI',sans-serif;box-shadow:0 2px 6px rgba(0,0,0,.18)";
    b.innerHTML="<span>Vous agissez en tant que <b>"+esc(d.nom||d.comme)+"</b> ("+esc(d.comme)+"). Vos enregistrements sont marqués « par "+esc(d.vrai_nom||d.vrai)+" ».</span>"+
      "<button type=\"button\" style=\"font:600 12.5px system-ui,sans-serif;border:1.5px solid #fff;background:#fff;color:#8C3A1C;border-radius:99px;padding:5px 12px;cursor:pointer\">Revenir à mon compte</button>";
    b.querySelector("button").onclick=function(){this.disabled=true;rpc("quitter_role").then(function(){location.href="/";}).catch(function(e){b.querySelector("span").textContent="Impossible de revenir à votre compte : "+e.message;});};
    document.body.insertBefore(b,document.body.firstChild);}
  function verif(essai){rpc("role_actuel").then(function(d){if(d&&d.comme)bandeau(d);}).catch(function(){if(essai<2)setTimeout(function(){verif(essai+1);},2500);});}
  window.KepComme={prendre:function(email){return rpc("prendre_role",{p_email:email});},quitter:function(){return rpc("quitter_role");}};
  if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",function(){verif(0);});else verif(0);
})();
