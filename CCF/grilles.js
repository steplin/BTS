/* Moteur de publipostage des grilles : travaille directement sur le XML du classeur
   pour conserver la mise en page d'origine. Dépend de JSZip, DOMParser, XMLSerializer. */
(function(root){
const NS="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RNS="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PNS="http://schemas.openxmlformats.org/package/2006/relationships";
const CTNS="http://schemas.openxmlformats.org/package/2006/content-types";
const WS_TYPE="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const WS_CT="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

function env(){return {DP:root.DOMParser,XS:root.XMLSerializer};}
function parse(s){return new (env().DP)().parseFromString(s,"application/xml");}
function ser(d){return new (env().XS)().serializeToString(d);}
function kids(el,name){const o=[];if(!el)return o;for(let n=el.firstChild;n;n=n.nextSibling)if(n.nodeType===1&&(n.localName||n.nodeName.split(":").pop())===name)o.push(n);return o;}
function kid(el,name){return kids(el,name)[0]||null;}
function all(doc,name){return Array.from(doc.getElementsByTagNameNS("*",name));}
function colNum(c){let n=0;for(const ch of c)n=n*26+ch.charCodeAt(0)-64;return n;}
function colStr(n){let s="";while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
function splitRef(r){const m=/^([A-Z]+)(\d+)$/.exec(r);return m?{c:colNum(m[1]),r:+m[2]}:null;}
function dirOf(p){return p.includes("/")?p.slice(0,p.lastIndexOf("/")+1):"";}
function resolve(base,target){if(target.startsWith("/"))return target.slice(1);const parts=(dirOf(base)+target).split("/");const out=[];for(const p of parts){if(p==="..")out.pop();else if(p!==".")out.push(p);}return out.join("/");}
function relsPathOf(p){return dirOf(p)+"_rels/"+p.slice(dirOf(p).length)+".rels";}
function relTarget(fromPart,toPart){const a=dirOf(fromPart).split("/").filter(Boolean),b=toPart.split("/");let i=0;while(i<a.length&&a[i]===b[i])i++;return "../".repeat(a.length-i)+b.slice(i).join("/");}
function xmlEsc(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}

class Sheet{
  constructor(doc,sst){this.doc=doc;this.sst=sst;this.sd=all(doc,"sheetData")[0];}
  rows(){return kids(this.sd,"row");}
  row(r,create){for(const x of this.rows()){const n=+x.getAttribute("r");if(n===r)return x;if(n>r&&create){const e=this.doc.createElementNS(NS,"row");e.setAttribute("r",r);this.sd.insertBefore(e,x);return e;}}
    if(!create)return null;const e=this.doc.createElementNS(NS,"row");e.setAttribute("r",r);this.sd.appendChild(e);return e;}
  cell(ref,create){const p=splitRef(ref);const row=this.row(p.r,create);if(!row)return null;
    for(const c of kids(row,"c")){const q=splitRef(c.getAttribute("r"));if(q.c===p.c)return c;if(q.c>p.c&&create){const e=this.doc.createElementNS(NS,"c");e.setAttribute("r",ref);row.insertBefore(e,c);return e;}}
    if(!create)return null;const e=this.doc.createElementNS(NS,"c");e.setAttribute("r",ref);row.appendChild(e);return e;}
  text(ref){const c=typeof ref==="string"?this.cell(ref,false):ref;if(!c)return "";const t=c.getAttribute("t");
    if(t==="s"){const v=kid(c,"v");return v?(this.sst[+v.textContent]||""):"";}
    if(t==="inlineStr"){return all(c,"t").map(x=>x.textContent).join("");}
    const v=kid(c,"v");return v?v.textContent:"";}
  clear(c){while(c.firstChild)c.removeChild(c.firstChild);c.removeAttribute("t");}
  setStr(ref,s){const c=this.cell(ref,true);this.clear(c);c.setAttribute("t","inlineStr");const is=this.doc.createElementNS(NS,"is"),t=this.doc.createElementNS(NS,"t");t.setAttribute("xml:space","preserve");t.textContent=s;is.appendChild(t);c.appendChild(is);}
  setNum(ref,n){const c=this.cell(ref,true);this.clear(c);c.setAttribute("t","n");const v=this.doc.createElementNS(NS,"v");v.textContent=String(n);c.appendChild(v);}
  setFormula(ref,f,cachedStr){const c=this.cell(ref,true);this.clear(c);const fe=this.doc.createElementNS(NS,"f");fe.textContent=f;c.appendChild(fe);
    if(cachedStr!=null){c.setAttribute("t","str");const v=this.doc.createElementNS(NS,"v");v.textContent=cachedStr;c.appendChild(v);}}
  find(pred){for(const row of this.rows())for(const c of kids(row,"c")){const t=this.text(c);if(pred(t,c))return {cell:c,ref:c.getAttribute("r"),...splitRef(c.getAttribute("r")),text:t};}return null;}
  rowCells(r){const row=this.row(r,false);return row?kids(row,"c").map(c=>({c,ref:c.getAttribute("r"),col:splitRef(c.getAttribute("r")).c,text:this.text(c)})):[];}
  maxCol(){let m=1;for(const row of this.rows())for(const c of kids(row,"c"))m=Math.max(m,splitRef(c.getAttribute("r")).c);return m;}
}

async function readSST(zip,wbPath,wbRels){
  const rel=wbRels.find(r=>/sharedStrings$/.test(r.type));if(!rel)return [];
  const x=await zip.file(resolve(wbPath,rel.target))?.async("string");if(!x)return [];
  return all(parse(x),"si").map(si=>{const ts=[];(function walk(n){for(let k=n.firstChild;k;k=k.nextSibling){if(k.nodeType!==1)continue;const ln=k.localName;if(ln==="rPh")continue;if(ln==="t")ts.push(k.textContent);else walk(k);}})(si);return ts.join("");});
}
function readRels(xml){if(!xml)return [];return all(parse(xml),"Relationship").map(r=>({id:r.getAttribute("Id"),type:r.getAttribute("Type"),target:r.getAttribute("Target"),mode:r.getAttribute("TargetMode")}));}

async function openBook(buf){
  const zip=await root.JSZip.loadAsync(buf);
  const rootRels=readRels(await zip.file("_rels/.rels").async("string"));
  const wbPath=resolve("",rootRels.find(r=>/officeDocument$/.test(r.type)).target);
  const wbRelsPath=relsPathOf(wbPath);
  const wbXml=await zip.file(wbPath).async("string");
  const wbRels=readRels(await zip.file(wbRelsPath).async("string"));
  const sst=await readSST(zip,wbPath,wbRels);
  const wbDoc=parse(wbXml);
  const sheets=all(wbDoc,"sheet").map(s=>{const rid=s.getAttributeNS(RNS,"id")||s.getAttribute("r:id");const rel=wbRels.find(r=>r.id===rid);return {el:s,name:s.getAttribute("name"),rid,path:resolve(wbPath,rel.target)};});
  return {zip,wbPath,wbRelsPath,wbDoc,wbRels,sst,sheets};
}

/* Analyse d'un classeur vierge */
async function inspect(buf,fileName){
  const b=await openBook(buf);
  const notes=b.sheets.find(s=>/^notes/i.test(s.name.trim()));
  const model=b.sheets.find(s=>/mod[eè]le/i.test(s.name))||b.sheets.find(s=>s!==notes);
  if(!notes||!model)throw new Error("onglets Notes / Modèle introuvables");
  const code=((notes.name.match(/C\s?\d\.\d/i)||model.name.match(/C\s?\d\.\d/i)||fileName.match(/C\s?\d\.\d/i)||[""])[0]).replace(/\s/g,"").toUpperCase();
  const ns=new Sheet(parse(await b.zip.file(notes.path).async("string")),b.sst);
  const lab=(re)=>{const f=ns.find(t=>re.test(t));if(!f)return null;const next=ns.rowCells(f.r).find(x=>x.col>f.c&&x.text.trim());return next?{ref:next.ref,value:next.text.trim()}:{ref:null,value:""};};
  const cap=lab(/^\s*capacit/i),se=lab(/situation/i),inter=lab(/intervenant/i),promo=lab(/promotion/i);
  const ms=new Sheet(parse(await b.zip.file(model.path).async("string")),b.sst);
  return {code,fileName,notesName:notes.name,modelName:model.name,
    titre:cap?cap.value.replace(/^C\s?\d\.\d\s*/i,""):"",se:se?se.value:"",intervenant:inter?inter.value:"",promo:promo?promo.value:"",
    ...gridStructure(ms)};
}
/* Lignes critères / indicateurs / barèmes de l'onglet Modèle */
function gridStructure(ms){
  const h=ms.find(t=>/^\s*indicateurs\s*$/i.test(t));const tot=ms.find(t=>/^\s*total\b/i.test(t));
  if(!h||!tot)return {criteres:[],sur:20};
  const criteres=[];
  for(let r=h.r+1;r<tot.r;r++){const B=ms.text("B"+r).trim(),C=ms.text("C"+r),I=ms.text("I"+r);
    if(!C.trim()&&!B)continue;
    if(B||!criteres.length)criteres.push({c:B,row:r,i:[]});
    criteres[criteres.length-1].i.push({t:C.replace(/^\n+/,"").replace(/\s+$/,""),b:+String(I).replace(",",".")||0,row:r});}
  const sur=+ms.text("I"+tot.r)||20;
  return {criteres,sur,totalRow:tot.r};
}

function titleCase(s){return s.toLowerCase().replace(/(^|[\s\-'’])(\p{L})/gu,(m,a,b)=>a+b.toUpperCase());}
function safeSheet(s){return s.replace(/[\[\]:*?\/\\]/g," ").replace(/\s+/g," ").trim();}
function q(n){return "'"+n.replace(/'/g,"''")+"'";}

/* Génération d'un classeur rempli */
async function fill(buf,opts){
  const {apprenants,promo,se,intervenant,titre,date,edits,theme,structure,rename,guards}=opts;
  const F=Object.assign({dateFr:true,vcenter:true,logoRatio:true,unbox:true,logoFit:true,themes:true},opts.features||{});
  if(rename)buf=await renameCode(buf,rename.from,rename.to);
  const b=await openBook(buf);const {zip}=b;
  if(F.logoRatio)await fixDrawings(zip);
  const notesS=b.sheets.find(s=>/^notes/i.test(s.name.trim()));
  const modelS=b.sheets.find(s=>/mod[eè]le/i.test(s.name))||b.sheets.find(s=>s!==notesS);
  const code=((notesS.name.match(/C\s?\d\.\d/i)||modelS.name.match(/C\s?\d\.\d/i)||[""])[0]).replace(/\s/g,"").toUpperCase();
  const ns=new Sheet(parse(await zip.file(notesS.path).async("string")),b.sst);
  const modelXml=await zip.file(modelS.path).async("string");
  const ms=new Sheet(parse(modelXml),b.sst);

  /* En-têtes éditables de l'onglet Notes */
  const setLabel=(re,val)=>{if(val==null)return;const f=ns.find(t=>re.test(t));if(!f)return;const next=ns.rowCells(f.r).find(x=>x.col>f.c&&x.text.trim());
    if(next){if(kid(next.c,"f"))return;ns.setStr(next.ref,val);}};
  setLabel(/promotion/i,promo);setLabel(/situation/i,se);setLabel(/intervenant/i,intervenant);
  if(titre!=null){const f=ns.find(t=>/^\s*capacit/i.test(t));if(f){const next=ns.rowCells(f.r).find(x=>x.col>f.c&&x.text.trim());ns.setStr(next?next.ref:colStr(f.c+1)+f.r,code+" "+titre);}}
  if(date){const f=ns.find(t=>/^\s*date\s*$/i.test(t));if(f){const next=ns.rowCells(f.r).find(x=>x.col>f.c&&(x.text.trim()||kid(x.c,"f")));
    if(next){const [y,m,d]=date.split("-").map(Number);const serial=Math.round((Date.UTC(y,m-1,d)-Date.UTC(1899,11,30))/864e5);ns.setNum(next.ref,serial);}}}
  /* Textes et barèmes modifiés dans la page */
  if(edits&&edits.length){let changedB=false;
    for(const e of edits){if(e.c!=null)ms.setStr("B"+e.row,e.c);if(e.t!=null)ms.setStr("C"+e.row,e.t);if(e.b!=null){ms.setNum("I"+e.row,e.b);changedB=true;}}
    if(changedB){const st=gridStructure(ms);const sum=st.criteres.reduce((a,c)=>a+c.i.reduce((x,i)=>x+i.b,0),0);if(st.totalRow)ms.setNum("I"+st.totalRow,sum);}}
  /* Grille restructurée : critères et indicateurs ajoutés, supprimés ou nouvelle grille */
  if(structure){const capTxt=(rename?rename.to:code)+" "+(titre!=null?titre:"");
    const r=restructure(ms,structure,titre!=null?capTxt:null);
    if(r.delta){const mIdx=all(b.wbDoc,"sheet").indexOf(modelS.el);
      for(const d of all(b.wbDoc,"definedName"))if(d.getAttribute("localSheetId")===String(mIdx))d.textContent=d.textContent.replace(/\$?([A-Z]{1,3})\$(\d+)/g,(m,c,n)=>+n>=r.from?"$"+c+"$"+(+n+r.delta):m);}
    const sur=structure.sur;if(sur){const nh=ns.find(t=>/^\s*notes?\s*\/\s*\d+/i.test(t));if(nh)ns.setStr(nh.ref,ns.text(nh.ref).replace(/\d+/,String(sur)));}}
  if(F.logoFit){const pr=ms.find(t=>/^\s*promotion/i.test(t));if(pr&&pr.r>1)await fitLogos(zip,modelS.path,ms,pr.r-1);
    const si=ns.find(t=>/situation/i.test(t));if(si&&si.r>1)await fitLogos(zip,notesS.path,ns,si.r-1);}
  /* valeurs en cache des formules de la grille retirées : recalculées à l'ouverture (titres, totaux à jour partout) */
  for(const f of all(ms.doc,"f")){const c=f.parentNode;const v=kid(c,"v");if(v)c.removeChild(v);if(c.getAttribute("t")==="str")c.removeAttribute("t");}
  await styleBook(b,ms,ns,F.themes?(theme||"origine"):"origine",F);
  if(guards&&(guards.notes||guards.lock))await addGuards(b,ms,guards);
  const modelXmlEd=ser(ms.doc);

  /* Zone des apprenants */
  const head=ns.find(t=>/^\s*nom\s*&\s*pr/i.test(t)&&!/intervenant/i.test(t));if(!head)throw new Error("ligne « Nom & Prénom » introuvable dans "+notesS.name);
  let start=head.r+1;
  if(ns.rowCells(start).some(x=>/^\s*\/\s*\d+/.test(x.text)))start++;
  const moy=ns.find(t=>/^\s*moyenne/i.test(t));
  /* Total(s) de la grille modèle */
  const totAll=ms.find(t=>/^\s*total\b/i.test(t)&&!/sur\s+20\b/i.test(t));
  const t20=ms.find(t=>/^\s*total\s+sur\s+20\b/i.test(t));
  const hRef=row=>{const cells=ms.rowCells(row).filter(x=>x.col>1&&kid(x.c,"f"));return (cells.length?colStr(cells[0].col):"H")+row;};
  // colonnes de l'onglet Notes à relier
  const links=[];
  const hdr=ns.rowCells(head.r);const ab=hdr.find(x=>/a\s*\+\s*b/i.test(x.text));
  if(ab){const sub=ns.rowCells(head.r+1).filter(x=>x.col>=ab.col);
    const c200=sub.find(x=>/\/\s*200\b/.test(x.text)),c20=sub.find(x=>/\/\s*20\b/.test(x.text));
    if(c200&&totAll)links.push({col:c200.col,ref:hRef(totAll.r)});if(c20&&(t20||totAll))links.push({col:c20.col,ref:hRef((t20||totAll).r)});}
  else{const n=hdr.find(x=>/note/i.test(x.text));const tt=t20||totAll;if(tt)links.push({col:n?n.col:3,ref:hRef(tt.r)});}
  const N=apprenants.length;
  let capacity=moy?moy.r-start:Infinity;
  if(moy&&N>capacity){shiftRows(ns,moy.r,N-capacity,start,moy.r-1);capacity=N;}
  // ligne type pour styles
  const tmplRow=ns.row(start,false)||ns.row(start-1,false);
  const ensureRow=r=>{let row=ns.row(r,false);if(row)return row;row=ns.row(r,true);
    if(tmplRow){for(const a of ["s","customFormat","ht","customHeight"])if(tmplRow.getAttribute(a)!=null)row.setAttribute(a,tmplRow.getAttribute(a));
      for(const c of kids(tmplRow,"c")){const col=splitRef(c.getAttribute("r")).c;const e=ns.doc.createElementNS(NS,"c");e.setAttribute("r",colStr(col)+r);if(c.getAttribute("s"))e.setAttribute("s",c.getAttribute("s"));row.appendChild(e);}
      const tr=+tmplRow.getAttribute("r");const mc=all(ns.doc,"mergeCells")[0];
      if(mc)for(const m of kids(mc,"mergeCell")){const [a,b2]=m.getAttribute("ref").split(":");const pa=splitRef(a),pb=splitRef(b2||a);
        if(pa.r===tr&&pb.r===tr){const e=ns.doc.createElementNS(NS,"mergeCell");e.setAttribute("ref",colStr(pa.c)+r+":"+colStr(pb.c)+r);mc.appendChild(e);mc.setAttribute("count",kids(mc,"mergeCell").length);}}
    }
    return row;};

  /* Onglets par apprenant */
  const used=new Set(b.sheets.filter(s=>s!==modelS).map(s=>s.name.toLowerCase()));
  const sheetNames=apprenants.map(a=>{let base=safeSheet(code+" "+titleCase(a.nom)+" "+titleCase(a.prenom)).slice(0,31),s=base,k=2;
    while(used.has(s.toLowerCase())){const suf=" "+k++;s=base.slice(0,31-suf.length)+suf;}used.add(s.toLowerCase());return s;});

  apprenants.forEach((a,i)=>{const r=start+i;ensureRow(r);
    ns.setStr("A"+r,a.nom.toUpperCase());ns.setStr("B"+r,titleCase(a.prenom));
    for(const l of links)ns.setFormula(colStr(l.col)+r,q(sheetNames[i])+"!"+l.ref);});

  if(!N){return {name:null,data:await zip.generateAsync({type:"uint8array"})};}

  // contenu de chaque grille
  const nameCellRef=(()=>{const f=ms.find(t=>/^\s*nom\s*$/i.test(t));if(!f)return {nom:"H2",prenom:"H3"};
    const p=ms.find(t=>/^\s*pr[ée]nom\s*$/i.test(t));const target=Math.max(f.c+1,colNum("H"));return {nom:colStr(target)+f.r,prenom:colStr(target)+(p?p.r:f.r+1)};})();

  const ct=parse(await zip.file("[Content_Types].xml").async("string"));const ctRoot=ct.documentElement;
  const overrides=()=>all(ct,"Override");
  const addOverride=(part,type)=>{const o=ct.createElementNS(CTNS,"Override");o.setAttribute("PartName","/"+part);o.setAttribute("ContentType",type);ctRoot.appendChild(o);};
  const ctOf=part=>{const o=overrides().find(x=>x.getAttribute("PartName")==="/"+part);return o?o.getAttribute("ContentType"):null;};

  const wbRelsDoc=parse(await zip.file(b.wbRelsPath).async("string"));
  const relRoot=wbRelsDoc.documentElement;
  let ridN=1;const rids=new Set(all(wbRelsDoc,"Relationship").map(r=>r.getAttribute("Id")));const newRid=()=>{while(rids.has("rId"+ridN))ridN++;const id="rId"+ridN;rids.add(id);return id;};
  const sheetsEl=all(b.wbDoc,"sheets")[0];
  let sheetId=Math.max(...all(b.wbDoc,"sheet").map(s=>+s.getAttribute("sheetId")))+1;
  const existing=new Set(Object.keys(zip.files));
  const uniquePart=(p)=>{const m=/^(.*?)(\d*)(\.[^.\/]+)$/.exec(p);let n=1,out;do{out=m[1].replace(/\d+$/,"")+(100+n++)+m[3];}while(existing.has(out));existing.add(out);return out;};
  const modelRelsPath=relsPathOf(modelS.path);
  const modelRelsXml=zip.file(modelRelsPath)?await zip.file(modelRelsPath).async("string"):null;

  // noms définis (zone d'impression…) du modèle
  const modelIdx=all(b.wbDoc,"sheet").indexOf(modelS.el);
  const dnEl=all(b.wbDoc,"definedNames")[0];
  const modelDefined=dnEl?kids(dnEl,"definedName").filter(d=>d.getAttribute("localSheetId")===String(modelIdx)):[];
  const sheetsAfter=all(b.wbDoc,"sheet").length-1-modelIdx;
  if(sheetsAfter>0&&dnEl){for(const d of kids(dnEl,"definedName")){const l=d.getAttribute("localSheetId");if(l!=null&&+l>modelIdx)d.setAttribute("localSheetId",+l+N-1);}}

  async function dupRels(relsXml,newOwner,oldOwner){
    if(!relsXml)return;const d=parse(relsXml);
    for(const r of all(d,"Relationship")){if(r.getAttribute("TargetMode")==="External")continue;
      const type=r.getAttribute("Type");if(!/(printerSettings|drawing|vmlDrawing|comments|threadedComment)$/i.test(type))continue;
      const oldPart=resolve(oldOwner,r.getAttribute("Target"));const f=zip.file(oldPart);if(!f)continue;
      const np=uniquePart(oldPart);zip.file(np,await f.async("uint8array"));
      const t=ctOf(oldPart);if(t)addOverride(np,t);
      const sub=relsPathOf(oldPart);if(zip.file(sub))await dupRels(await zip.file(sub).async("string"),np,oldPart).then(x=>{});
      r.setAttribute("Target",relTarget(newOwner,np));}
    zip.file(relsPathOf(newOwner),ser(d));}

  let insertAfter=modelS.el;
  /* grilles remplies (saisie en ligne) : lignes des indicateurs de la grille finale */
  const ROWS=gridStructure(new Sheet(parse(modelXmlEd),b.sst)).criteres.flatMap(c=>c.i.map(x=>x.row));
  const lvOf=(n,bb)=>{if(!(bb>0))return -1;const r=n/bb;return r<0.25?0:r<0.5?1:r<0.75?2:3;};
  for(let i=0;i<N;i++){
    const nm=sheetNames[i];let path;
    if(i===0){path=modelS.path;modelS.el.setAttribute("name",nm);}
    else{path=uniquePart(modelS.path);addOverride(path,WS_CT);
      const rid=newRid();const rel=wbRelsDoc.createElementNS(PNS,"Relationship");rel.setAttribute("Id",rid);rel.setAttribute("Type",WS_TYPE);rel.setAttribute("Target",relTarget(b.wbPath,path));relRoot.appendChild(rel);
      const se=b.wbDoc.createElementNS(NS,"sheet");se.setAttribute("name",nm);se.setAttribute("sheetId",sheetId++);se.setAttributeNS(RNS,"r:id",rid);
      sheetsEl.insertBefore(se,insertAfter.nextSibling);insertAfter=se;
      await dupRels(modelRelsXml,path,modelS.path);}
    const s=new Sheet(parse(modelXmlEd),b.sst);
    const r=start+i;
    s.setFormula(nameCellRef.nom,q(notesS.name)+"!A"+r,apprenants[i].nom.toUpperCase());
    s.setFormula(nameCellRef.prenom,q(notesS.name)+"!B"+r,titleCase(apprenants[i].prenom));
    const sz=apprenants[i].saisie;
    if(sz&&Array.isArray(sz.lignes)){sz.lignes.forEach((l,k)=>{const rr=ROWS[k];if(rr==null)return;
        if(l.n!=null&&l.n!==""){s.setNum("H"+rr,+l.n);const lv=lvOf(+l.n,+l.b);if(lv>=0)s.setStr("DEFG"[lv]+rr,"X");}
        if(l.a)s.setStr("J"+rr,l.a);});
      if(sz.ag){const ap=s.find(t=>/appr[ée]ciation\s+g[ée]n[ée]rale/i.test(t));if(ap)s.setStr(ap.ref,s.text(ap.ref).replace(/\s+$/,"")+" "+sz.ag);}}
    for(const sv of all(s.doc,"sheetView"))sv.removeAttribute("tabSelected");
    zip.file(path,ser(s.doc));
    if(i>0)for(const d of modelDefined){const c=d.cloneNode(true);c.setAttribute("localSheetId",modelIdx+i);c.textContent=d.textContent.split(q(modelS.name)).join(q(nm)).split(modelS.name+"!").join(q(nm)+"!");dnEl.appendChild(c);}
  }
  for(const d of modelDefined)d.textContent=d.textContent.split(q(modelS.name)).join(q(sheetNames[0])).split(modelS.name+"!").join(q(sheetNames[0])+"!");
  if(dnEl){const arr=kids(dnEl,"definedName").sort((a,b)=>(a.getAttribute("name")+"|"+(a.getAttribute("localSheetId")||"")).localeCompare(b.getAttribute("name")+"|"+(b.getAttribute("localSheetId")||""),"en",{numeric:true}));arr.forEach(x=>dnEl.appendChild(x));}

  // Notes : onglet actif
  for(const sv of all(ns.doc,"sheetView"))sv.setAttribute("tabSelected","1");
  const notesIdx=all(b.wbDoc,"sheet").findIndex(s=>s.getAttribute("name")===notesS.name);
  for(const v of all(b.wbDoc,"workbookView")){v.setAttribute("activeTab",notesIdx);v.removeAttribute("firstSheet");}
  zip.file(notesS.path,ser(ns.doc));

  // recalcul complet à l'ouverture, suppression de la chaîne de calcul
  let calc=all(b.wbDoc,"calcPr")[0];if(!calc){calc=b.wbDoc.createElementNS(NS,"calcPr");b.wbDoc.documentElement.appendChild(calc);}calc.setAttribute("fullCalcOnLoad","1");
  for(const r of all(wbRelsDoc,"Relationship"))if(/calcChain$/.test(r.getAttribute("Type"))){const p=resolve(b.wbPath,r.getAttribute("Target"));zip.remove(p);r.parentNode.removeChild(r);for(const o of overrides())if(o.getAttribute("PartName")==="/"+p)o.parentNode.removeChild(o);}
  zip.file(b.wbPath,ser(b.wbDoc));zip.file(b.wbRelsPath,ser(wbRelsDoc));zip.file("[Content_Types].xml",ser(ct));
  // propriétés : la liste des onglets n'est plus exacte, on la retire
  const app=zip.file("docProps/app.xml");if(app){let x=await app.async("string");x=x.replace(/<HeadingPairs>[\s\S]*?<\/HeadingPairs>/,"").replace(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/,"");zip.file("docProps/app.xml",x);}
  return {data:await zip.generateAsync({type:"uint8array",compression:"DEFLATE"})};
}

function shiftRows(sh,from,k,firstData,lastData){
  // décale les lignes >= from de k lignes (ligne Moyenne et suivantes)
  const rows=sh.rows().filter(r=>+r.getAttribute("r")>=from);
  const bump=s=>s.replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g,(m,c,n)=>+n>=from?c+(+n+k):(lastData!=null&&+n===lastData?c+(lastData+k):m));
  for(const row of (k>0?rows.reverse():rows)){const r=+row.getAttribute("r");row.setAttribute("r",r+k);
    for(const c of kids(row,"c")){const p=splitRef(c.getAttribute("r"));c.setAttribute("r",colStr(p.c)+(r+k));const f=kid(c,"f");if(f){f.textContent=bump(f.textContent);if(f.getAttribute("ref"))f.setAttribute("ref",bump(f.getAttribute("ref")));}}}
  for(const m of all(sh.doc,"mergeCell")){const ref=m.getAttribute("ref");const [a,b]=ref.split(":");const pa=splitRef(a),pb=splitRef(b||a);if(pa.r>=from)m.setAttribute("ref",bump(ref));}
  for(const tag of ["conditionalFormatting","dataValidation"])for(const e of all(sh.doc,tag)){const s=e.getAttribute("sqref");if(s)e.setAttribute("sqref",s.split(" ").map(x=>{const p=x.split(":").map(splitRef);return p[0]&&p[0].r>=from?bump(x):x;}).join(" "));}
  const dim=all(sh.doc,"dimension")[0];if(dim)dim.parentNode.removeChild(dim);
}


/* ---------- Grille restructurée ---------- */
/* Reconstruit la zone critères/indicateurs de l'onglet Modèle : st = {criteres:[{c,i:[{t,b}]}]} */
function restructure(ms,st,capText){
  const h=ms.find(t=>/^\s*indicateurs\s*$/i.test(t));const tot=ms.find(t=>/^\s*total\b/i.test(t));
  if(!h||!tot)throw new Error("zone des critères introuvable dans l'onglet Modèle");
  const s0=h.r+1,e0=tot.r-1,n0=e0-s0+1;
  const rowsIn=st.criteres.flatMap(c=>c.i.map((it,k)=>({c:k===0?c.c:null,t:it.t,b:+it.b||0,first:k===0,n:c.i.length})));
  const n1=rowsIn.length;if(!n1)throw new Error("la grille doit contenir au moins un indicateur");
  const tFirst=ms.row(s0,false),tMid=ms.row(Math.min(s0+1,e0),false)||tFirst;
  const capCell=ms.cell("A"+s0,false);const capClone=capCell?capCell.cloneNode(true):null;
  // largeurs de colonnes (pour estimer les hauteurs de ligne)
  const widths={};for(const c of all(ms.doc,"col")){for(let k=+c.getAttribute("min");k<=+c.getAttribute("max");k++)widths[k]=+c.getAttribute("width");}
  const fmt=all(ms.doc,"sheetFormatPr")[0];const defH=+(fmt&&fmt.getAttribute("defaultRowHeight"))||15;
  const origHts=[];for(let r=s0;r<=e0;r++){const x=ms.row(r,false);if(x&&x.getAttribute("ht"))origHts.push(+x.getAttribute("ht"));}
  const minH=origHts.length?Math.min(...origHts):defH;
  const lines=(txt,col)=>{const w=widths[col]||10;const cpl=Math.max(8,Math.floor(w*1.15));return String(txt||"").split("\n").reduce((a,l)=>a+Math.max(1,Math.ceil(l.length/cpl)),0);};
  // fusions de la zone retirées
  const mc=all(ms.doc,"mergeCells")[0];
  if(mc){for(const m of kids(mc,"mergeCell")){const [a,b2]=m.getAttribute("ref").split(":");const pa=splitRef(a),pb=splitRef(b2||a);if(pb.r>=s0&&pa.r<=e0)mc.removeChild(m);}}
  // lignes de la zone retirées, lignes suivantes décalées
  for(let r=s0;r<=e0;r++){const x=ms.row(r,false);if(x)x.parentNode.removeChild(x);}
  const delta=n1-n0;const from=tot.r;
  if(delta)shiftRows(ms,from,delta);
  // nouvelles lignes
  const mk=(tpl,r)=>{const row=ms.doc.createElementNS(NS,"row");row.setAttribute("r",r);
    if(tpl){for(const a of ["s","customFormat","spans","x14ac:dyDescent"])if(tpl.getAttribute(a)!=null)row.setAttribute(a,tpl.getAttribute(a));
      for(const c of kids(tpl,"c")){const col=splitRef(c.getAttribute("r")).c;const e=ms.doc.createElementNS(NS,"c");e.setAttribute("r",colStr(col)+r);if(c.getAttribute("s"))e.setAttribute("s",c.getAttribute("s"));row.appendChild(e);}}
    let after=null;for(const x of ms.rows()){if(+x.getAttribute("r")>r){after=x;break;}}ms.sd.insertBefore(row,after);return row;};
  rowsIn.forEach((x,k)=>{const r=s0+k;const row=mk(k===0?tFirst:tMid,r);
    if(x.first&&x.c!=null)ms.setStr("B"+r,x.c);ms.setStr("C"+r,x.t);ms.setNum("I"+r,x.b);
    const need=Math.max(lines(x.t,3),x.first&&x.n===1?lines(x.c,2):Math.ceil(lines(x.c||"",2)/x.n));
    const ht=Math.max(minH,Math.round(need*14.4+5));row.setAttribute("ht",ht);row.setAttribute("customHeight","1");});
  // capacité en colonne A
  if(capClone){const c=ms.cell("A"+s0,true);while(c.firstChild)c.removeChild(c.firstChild);for(const a of Array.from(capClone.attributes))if(a.name!=="r")c.setAttribute(a.name,a.value);for(let n=capClone.firstChild;n;n=n.nextSibling)c.appendChild(n.cloneNode(true));}
  if(capText)ms.setStr("A"+s0,capText);
  // fusions : capacité sur toute la zone, chaque critère sur ses indicateurs
  let m2=mc;if(!m2){m2=ms.doc.createElementNS(NS,"mergeCells");const after=all(ms.doc,"sheetData")[0];after.parentNode.insertBefore(m2,after.nextSibling);}
  const addM=ref=>{const e=ms.doc.createElementNS(NS,"mergeCell");e.setAttribute("ref",ref);m2.appendChild(e);};
  if(n1>1)addM("A"+s0+":A"+(s0+n1-1));
  let r=s0;for(const c of st.criteres){if(c.i.length>1)addM("B"+r+":B"+(r+c.i.length-1));r+=c.i.length;}
  m2.setAttribute("count",kids(m2,"mergeCell").length);
  // totaux : sommes recalées sur la nouvelle zone
  const s1=s0,e1=s0+n1-1,tr=tot.r+delta;
  for(const row of ms.rows()){if(+row.getAttribute("r")<tr)continue;for(const c of kids(row,"c")){const f=kid(c,"f");if(!f)continue;
    f.textContent=f.textContent.replace(/SUM\((\$?[A-Z]{1,3})\$?(\d+):(\$?[A-Z]{1,3})\$?(\d+)\)/gi,(m,ca,ra,cb,rb)=>(+ra>=s0-1&&+ra<=s0&&+rb>=s0&&+rb<=e0)?"SUM("+ca+s1+":"+cb+e1+")":m);
    const v=kid(c,"v");if(v)c.removeChild(v);}}
  const iTot=ms.cell("I"+tr,false);if(iTot&&!kid(iTot,"f"))ms.setNum("I"+tr,rowsIn.reduce((a,x)=>a+x.b,0));
  const dim=all(ms.doc,"dimension")[0];if(dim)dim.parentNode.removeChild(dim);
  return {delta,from,first:s1,last:e1,totalRow:tr};
}
/* Renomme les onglets d'une grille de base (« Notes C1.1 », « C1.1 Modèle ») pour une nouvelle capacité */
async function renameCode(buf,from,to){
  const b=await openBook(buf);const {zip}=b;const esc=s=>s.replace(/[.*+?^${}()|[\]\\]/g,"\\$&");
  const map={};for(const s of b.sheets){const nn=s.name.replace(new RegExp(esc(from),"g"),to);if(nn!==s.name){map[s.name]=nn;s.el.setAttribute("name",nn);}}
  const fix=x=>{for(const [o,n] of Object.entries(map)){x=x.split(q(o)).join(q(n)).split(o+"!").join(q(n)+"!");}return x;};
  for(const s of b.sheets){const x=await zip.file(s.path).async("string");const d=parse(x);let ch=false;for(const f of all(d,"f")){const t=fix(f.textContent);if(t!==f.textContent){f.textContent=t;ch=true;}}if(ch)zip.file(s.path,ser(d));}
  for(const d of all(b.wbDoc,"definedName"))d.textContent=fix(d.textContent);
  zip.file(b.wbPath,ser(b.wbDoc));
  const app=zip.file("docProps/app.xml");if(app){let x=await app.async("string");for(const [o,n] of Object.entries(map))x=x.split(xmlEsc(o)).join(xmlEsc(n));zip.file("docProps/app.xml",x);}
  return await zip.generateAsync({type:"uint8array"});
}
/* ---------- Garde-fous : contrôle des notes et verrouillage des grilles ---------- */
const ORDER=["sheetPr","dimension","sheetViews","sheetFormatPr","cols","sheetData","sheetCalcPr","sheetProtection","protectedRanges","scenarios","autoFilter","sortState","dataConsolidate","customSheetViews","mergeCells","phoneticPr","conditionalFormatting","dataValidations","hyperlinks","printOptions","pageMargins","pageSetup","headerFooter","rowBreaks","colBreaks","customProperties","cellWatches","ignoredErrors","smartTags","drawing","legacyDrawing","legacyDrawingHF","drawingHF","picture","oleObjects","controls","webPublishItems","tableParts","extLst"];
function placeEl(doc,el){const rootEl=doc.documentElement;const i=ORDER.indexOf(el.localName);
  let before=null;for(let n=rootEl.firstChild;n;n=n.nextSibling){if(n.nodeType!==1)continue;const j=ORDER.indexOf(n.localName);if(j>i){before=n;break;}}
  rootEl.insertBefore(el,before);return el;}
function pwHash(pw){let h=0;for(let i=pw.length-1;i>=0;i--){h=((h>>14)&1)|((h<<1)&0x7fff);h^=pw.charCodeAt(i);}
  h=((h>>14)&1)|((h<<1)&0x7fff);h^=pw.length;h^=0xCE4B;return h.toString(16).toUpperCase().padStart(4,"0");}
async function addGuards(b,ms,g){
  const h=ms.find(t=>/^\s*indicateurs\s*$/i.test(t));const tot=ms.find(t=>/^\s*total\b/i.test(t));if(!h||!tot)return;
  const s0=h.r+1,e0=tot.r-1;const data=[];for(let r=s0;r<=e0;r++){if(String(ms.text("C"+r)).trim()||String(ms.text("I"+r)).trim())data.push(r);}
  const path=resolve(b.wbPath,(b.wbRels.find(r=>/styles$/.test(r.type))||{}).target||"styles.xml");
  const d=parse(await b.zip.file(path).async("string"));const cellXfs=all(d,"cellXfs")[0];
  const doc=ms.doc;
  if(g.notes&&data.length){
    // saisie refusée au-delà du barème (colonne I de la même ligne)
    let dvs=all(doc,"dataValidations")[0];if(!dvs)dvs=placeEl(doc,doc.createElementNS(NS,"dataValidations"));
    const dv=doc.createElementNS(NS,"dataValidation");
    const sq=data.map(r=>"H"+r).join(" ");
    for(const [k,v] of Object.entries({type:"decimal",allowBlank:"1",showInputMessage:"1",showErrorMessage:"1",errorStyle:"stop",
      errorTitle:"Note impossible",error:"La note doit être comprise entre 0 et le barème de la ligne (colonne I).",promptTitle:"Note",prompt:"Entre 0 et le barème de la ligne.",sqref:sq}))dv.setAttribute(k,v);
    const f1=doc.createElementNS(NS,"formula1");f1.textContent="0";const f2=doc.createElementNS(NS,"formula2");f2.textContent="$I"+data[0];dv.appendChild(f1);dv.appendChild(f2);
    dvs.appendChild(dv);dvs.setAttribute("count",kids(dvs,"dataValidation").length);
    // filet de sécurité (copier-coller) : note > barème en rouge, total > note maximale en rouge
    let dxfs=all(d,"dxfs")[0];if(!dxfs){dxfs=d.createElementNS(NS,"dxfs");const cs=all(d,"cellStyles")[0];cs.parentNode.insertBefore(dxfs,cs.nextSibling);}
    const dx=parse('<dxf xmlns="'+NS+'"><font><b/><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>').documentElement;
    dxfs.appendChild(d.importNode(dx,true));dxfs.setAttribute("count",kids(dxfs,"dxf").length);const dxfId=kids(dxfs,"dxf").length-1;
    let prio=100;const cf=(sqref,formula)=>{const c=placeEl(doc,doc.createElementNS(NS,"conditionalFormatting"));c.setAttribute("sqref",sqref);
      const rule=doc.createElementNS(NS,"cfRule");rule.setAttribute("type","expression");rule.setAttribute("dxfId",dxfId);rule.setAttribute("priority",prio++);
      const f=doc.createElementNS(NS,"formula");f.textContent=formula;rule.appendChild(f);c.appendChild(rule);};
    cf(sq,"AND(ISNUMBER(H"+data[0]+"),OR(H"+data[0]+">$I"+data[0]+",H"+data[0]+"<0))");
    const tH=ms.rowCells(tot.r).find(x=>x.col===8);if(tH)cf("H"+tot.r,"H"+tot.r+">I"+tot.r);
  }
  if(g.lock){
    // cellules de saisie déverrouillées : niveaux --/-/+/++, note, appréciation par ligne, appréciation générale
    const cache={};const unlocked=s=>{if(cache[s]!=null)return cache[s];const xf=kids(cellXfs,"xf")[s]||kids(cellXfs,"xf")[0];const n=xf.cloneNode(true);
      for(const p of kids(n,"protection"))n.removeChild(p);const p=d.createElementNS(NS,"protection");p.setAttribute("locked","0");n.appendChild(p);n.setAttribute("applyProtection","1");
      cellXfs.appendChild(n);cellXfs.setAttribute("count",kids(cellXfs,"xf").length);return cache[s]=kids(cellXfs,"xf").length-1;};
    const open=(r,c1,c2)=>{for(let c=c1;c<=c2;c++){const cell=ms.cell(colStr(c)+r,true);cell.setAttribute("s",unlocked(+cell.getAttribute("s")||0));}};
    for(const r of data){open(r,4,8);open(r,10,10);}
    const ap=ms.find(t=>/appr[ée]ciation\s+g[ée]n[ée]rale/i.test(t));
    if(ap){let last=ap.r;const maxR=Math.max(...ms.rows().map(x=>+x.getAttribute("r")));
      for(let r=ap.r;r<=Math.min(maxR,ap.r+6);r++){if(r>ap.r&&ms.rowCells(r).some(x=>x.text.trim()))break;last=r;}
      for(let r=ap.r;r<=last;r++)open(r,r===ap.r?Math.max(2,ap.c+1):1,10);
      // l'appréciation générale est souvent dans la cellule fusionnée contenant le libellé : on garde le libellé et on ouvre le reste
      open(ap.r,ap.c,ap.c);}
    let sp=all(doc,"sheetProtection")[0];if(!sp)sp=placeEl(doc,doc.createElementNS(NS,"sheetProtection"));
    for(const [k,v] of Object.entries({sheet:"1",objects:"1",scenarios:"1",formatRows:"0",formatColumns:"0",selectLockedCells:"0",selectUnlockedCells:"0"}))sp.setAttribute(k,v);
    if(g.password)sp.setAttribute("password",pwHash(g.password));
  }
  b.zip.file(path,ser(d));
}

/* ---------- Mise en forme : corrections communes et modèles de sortie ---------- */
const THEMES={
  origine:null,
  kerplouz:{font:"Calibri",header:{fill:"FF2E6A4B",color:"FFFFFFFF",b:true},soft:{fill:"FFE8F2EA",color:"FF1F4D36",b:true},
    total:{fill:"FFCFE3D4",color:"FF1F4D36",b:true},note:{fill:"FFFFF6D6"},label:{color:"FF2E6A4B",b:true}},
  impression:{font:"Arial",header:{fill:"FFE7E7E7",color:"FF000000",b:true},soft:{fill:"FFF5F5F5",color:"FF000000",b:true},
    total:{fill:"FFE7E7E7",color:"FF000000",b:true},note:{fill:"FFFFFFFF"},label:{color:"FF000000",b:true},size:11},
  logo:{font:"Century Gothic",header:{fill:"FFD1683F",color:"FFFFFFFF",b:true},soft:{fill:"FFEEF4E4",color:"FF5E6B1F",b:true},
    total:{fill:"FFDCEBC9",color:"FF2C2622",b:true},note:{fill:"FFFBE9E1"},label:{color:"FFD1683F",b:true}}
};
function imgSize(u8){
  if(u8[0]===0x89&&u8[1]===0x50){const v=new DataView(u8.buffer,u8.byteOffset);return {w:v.getUint32(16),h:v.getUint32(20)};}
  if(u8[0]===0xFF&&u8[1]===0xD8){let i=2;const v=new DataView(u8.buffer,u8.byteOffset);
    while(i<u8.length){if(u8[i]!==0xFF){i++;continue;}const m=u8[i+1];const len=v.getUint16(i+2);
      if(m>=0xC0&&m<=0xCF&&m!==0xC4&&m!==0xC8&&m!==0xCC)return {h:v.getUint16(i+5),w:v.getUint16(i+7)};i+=2+len;}}
  return null;
}
async function fixDrawings(zip){
  const X="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
  for(const p of Object.keys(zip.files).filter(p=>/^xl\/drawings\/drawing\d+\.xml$/.test(p))){
    const d=parse(await zip.file(p).async("string"));const rels=readRels(zip.file(relsPathOf(p))?await zip.file(relsPathOf(p)).async("string"):"");
    let changed=false;
    for(const anc of all(d,"twoCellAnchor")){
      const blip=all(anc,"blip")[0];const ext=all(anc,"ext").find(e=>e.parentNode&&e.parentNode.localName==="xfrm");if(!blip||!ext)continue;
      const rid=blip.getAttributeNS(RNS,"embed")||blip.getAttribute("r:embed");const rel=rels.find(r=>r.id===rid);if(!rel)continue;
      const f=zip.file(resolve(p,rel.target));if(!f)continue;const sz=imgSize(await f.async("uint8array"));if(!sz||!sz.w||!sz.h)continue;
      const cx=+ext.getAttribute("cx"),cy=+ext.getAttribute("cy");const ncx=Math.round(cy*sz.w/sz.h);
      const from=kid(anc,"from"),to=kid(anc,"to");
      // décalage pour garder le bord droit des logos placés à droite
      const fc=+kid(from,"col").textContent;if(fc>0){const co=kid(from,"colOff");co.textContent=Math.max(0,+co.textContent-(ncx-cx));}
      ext.setAttribute("cx",ncx);
      const one=d.createElementNS(X,"xdr:oneCellAnchor");one.appendChild(from.cloneNode(true));
      const e=d.createElementNS(X,"xdr:ext");e.setAttribute("cx",ncx);e.setAttribute("cy",cy);one.appendChild(e);
      for(let n=anc.firstChild;n;n=n.nextSibling){if(n!==from&&n!==to)one.appendChild(n.cloneNode(true));}
      anc.parentNode.replaceChild(one,anc);changed=true;}
    if(changed)zip.file(p,ser(d));}
}
/* Ajuste la hauteur des logos à la zone d'en-tête (lignes 1..lastRow) et les centre verticalement */
async function fitLogos(zip,sheetPath,sh,lastRow){
  const rp=relsPathOf(sheetPath);const f=zip.file(rp);if(!f)return;
  const rel=readRels(await f.async("string")).find(r=>/drawing$/.test(r.type));if(!rel)return;
  const dp=resolve(sheetPath,rel.target);const df=zip.file(dp);if(!df)return;
  const fmt=all(sh.doc,"sheetFormatPr")[0];const def=+(fmt&&fmt.getAttribute("defaultRowHeight"))||15;
  const ht=r=>{const row=sh.row(r,false);return (row&&row.getAttribute("ht")?+row.getAttribute("ht"):def)*12700;};
  let zone=0;for(let r=1;r<=lastRow;r++)zone+=ht(r);
  const d=parse(await df.async("string"));const margin=4*12700;
  for(const anc of [...all(d,"oneCellAnchor"),...all(d,"twoCellAnchor")]){
    const from=kid(anc,"from");if(!from)continue;const ext=kid(anc,"ext")||all(anc,"ext").find(e=>e.parentNode&&e.parentNode.localName==="xfrm");if(!ext)continue;
    let cx=+ext.getAttribute("cx"),cy=+ext.getAttribute("cy");const maxCy=zone-2*margin;
    if(cy>maxCy){const k=maxCy/cy;const ncx=Math.round(cx*k);
      if(+kid(from,"col").textContent>0){const co=kid(from,"colOff");co.textContent=Math.max(0,+co.textContent+(cx-ncx));}
      cx=ncx;cy=Math.round(maxCy);}
    // centrage vertical dans la zone
    let top=Math.max(margin,Math.round((zone-cy)/2)),r=0;while(r<lastRow-1&&top>=ht(r+1)){top-=ht(r+1);r++;}
    kid(from,"row").textContent=r;kid(from,"rowOff").textContent=top;
    for(const e of [ext,...all(anc,"ext").filter(e=>e.parentNode&&e.parentNode.localName==="xfrm")]){e.setAttribute("cx",cx);e.setAttribute("cy",cy);}
    if(anc.localName==="twoCellAnchor")anc.setAttribute("editAs","oneCell");}
  zip.file(dp,ser(d));
}
async function styleBook(b,ms,ns,themeName,F){F=F||{dateFr:true,vcenter:true,unbox:true};
  const path=resolve(b.wbPath,(b.wbRels.find(r=>/styles$/.test(r.type))||{}).target||"styles.xml");
  const f=b.zip.file(path);if(!f)return;const d=parse(await f.async("string"));
  // dates au format français
  if(F.dateFr)for(const n of all(d,"numFmt")){const c=n.getAttribute("formatCode")||"";if(/[md]\/[md]\/y/i.test(c)||/^m+\/d+\/y+$/i.test(c))n.setAttribute("formatCode","dd/mm/yyyy");}
  // toutes les cellules centrées verticalement
  const cellXfs=all(d,"cellXfs")[0];const xfs=kids(cellXfs,"xf");
  const center=xf=>{let a=kid(xf,"alignment");if(!a){a=d.createElementNS(NS,"alignment");xf.insertBefore(a,xf.firstChild);}a.setAttribute("vertical","center");xf.setAttribute("applyAlignment","true");};
  if(F.vcenter)xfs.forEach(center);
  /* ligne des logos : pas de bordure à gauche, à droite ni en haut, seulement en bas */
  const bordersEl=all(d,"borders")[0];const bCache={};
  const bVariant=(s,spec)=>{const key=s+"|"+JSON.stringify(spec);if(bCache[key]!=null)return bCache[key];
    const xf=kids(cellXfs,"xf")[s]||kids(cellXfs,"xf")[0];const base=kids(bordersEl,"border")[+xf.getAttribute("borderId")||0];
    const nb=base?base.cloneNode(true):d.createElementNS(NS,"border");
    for(const side of ["left","right","top","bottom"]){let e=kid(nb,side);
      if(!e){e=d.createElementNS(NS,side);const order=["left","right","top","bottom","diagonal"];const next=order.slice(order.indexOf(side)+1).map(x=>kid(nb,x)).find(Boolean);nb.insertBefore(e,next||null);}
      if(spec[side]===false){while(e.firstChild)e.removeChild(e.firstChild);e.removeAttribute("style");}
      else if(spec[side]===true&&!e.getAttribute("style")){e.setAttribute("style","thin");const c=d.createElementNS(NS,"color");c.setAttribute("auto","1");e.appendChild(c);}}
    bordersEl.appendChild(nb);bordersEl.setAttribute("count",kids(bordersEl,"border").length);
    const n=xf.cloneNode(true);n.setAttribute("borderId",kids(bordersEl,"border").length-1);n.setAttribute("applyBorder","true");
    cellXfs.appendChild(n);cellXfs.setAttribute("count",kids(cellXfs,"xf").length);return bCache[key]=kids(cellXfs,"xf").length-1;};
  const unbox=(sh,r1,r2,c1,c2)=>{for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++){const cell=sh.cell(colStr(c)+r,true);
    cell.setAttribute("s",bVariant(+cell.getAttribute("s")||0,{left:false,right:false,top:false,bottom:r===r2}));}};
  if(F.unbox){ const pr=ms.find(t=>/^\s*promotion/i.test(t));if(pr&&pr.r>1)unbox(ms,1,pr.r-1,1,ms.maxCol()); }
  if(F.unbox){ const dt=ns.find(t=>/^\s*date\s*$/i.test(t)),si=ns.find(t=>/situation/i.test(t));
    if(dt&&dt.c>1){const last=si&&si.r>dt.r?si.r-1:dt.r+1;unbox(ns,dt.r,last,1,dt.c-1);} }
  const T=THEMES[themeName];
  if(T){
    const fontsEl=all(d,"fonts")[0],fillsEl=all(d,"fills")[0];
    for(const fn of kids(fontsEl,"font")){const nm=kid(fn,"name");if(nm)nm.setAttribute("val",T.font);if(T.size){const s=kid(fn,"sz");if(s&&+s.getAttribute("val")<T.size)s.setAttribute("val",T.size);}}
    const fillCache={},fontCache={},xfCache={};
    const addFill=rgb=>{if(fillCache[rgb]!=null)return fillCache[rgb];const fl=d.createElementNS(NS,"fill"),pf=d.createElementNS(NS,"patternFill");pf.setAttribute("patternType","solid");
      const fg=d.createElementNS(NS,"fgColor");fg.setAttribute("rgb",rgb);const bg=d.createElementNS(NS,"bgColor");bg.setAttribute("indexed","64");pf.appendChild(fg);pf.appendChild(bg);fl.appendChild(pf);fillsEl.appendChild(fl);
      fillsEl.setAttribute("count",kids(fillsEl,"fill").length);return fillCache[rgb]=kids(fillsEl,"fill").length-1;};
    const addFont=(base,sp)=>{const key=base+"|"+JSON.stringify(sp);if(fontCache[key]!=null)return fontCache[key];const fn=kids(fontsEl,"font")[base].cloneNode(true);
      if(sp.b&&!kid(fn,"b")){const e=d.createElementNS(NS,"b");e.setAttribute("val","true");fn.insertBefore(e,fn.firstChild);}
      if(sp.color){let c=kid(fn,"color");if(!c){c=d.createElementNS(NS,"color");const sz=kid(fn,"sz");fn.insertBefore(c,sz?sz.nextSibling:fn.firstChild);}for(const a of ["theme","indexed","tint","auto"])c.removeAttribute(a);c.setAttribute("rgb",sp.color);}
      fontsEl.appendChild(fn);fontsEl.setAttribute("count",kids(fontsEl,"font").length);return fontCache[key]=kids(fontsEl,"font").length-1;};
    const variant=(s,role)=>{const sp=T[role];if(!sp)return s;const key=s+"|"+role;if(xfCache[key]!=null)return xfCache[key];
      const xf=kids(cellXfs,"xf")[s]||kids(cellXfs,"xf")[0];const n=xf.cloneNode(true);
      if(sp.fill){n.setAttribute("fillId",addFill(sp.fill));n.setAttribute("applyFill","true");}
      if(sp.color||sp.b){n.setAttribute("fontId",addFont(+xf.getAttribute("fontId")||0,{b:sp.b,color:sp.color}));n.setAttribute("applyFont","true");}
      cellXfs.appendChild(n);cellXfs.setAttribute("count",kids(cellXfs,"xf").length);return xfCache[key]=kids(cellXfs,"xf").length-1;};
    const paint=(sh,r,c1,c2,role)=>{const row=sh.row(r,false);if(!row)return;for(const c of kids(row,"c")){const col=splitRef(c.getAttribute("r")).c;if(col<c1||col>c2)continue;c.setAttribute("s",variant(+c.getAttribute("s")||0,role));}};
    // onglet Modèle
    const h=ms.find(t=>/^\s*indicateurs\s*$/i.test(t));
    if(h){const mc=ms.maxCol();paint(ms,h.r,1,mc,"header");
      for(let r=2;r<h.r;r++){for(const x of ms.rowCells(r))if(/^\s*(promotion|nom|pr[ée]nom)\s*$/i.test(x.text))paint(ms,r,x.col,x.col,"label");}
      let r=h.r+1;for(;;r++){const t=ms.text("A"+r);if(/^\s*total/i.test(t)||r>h.r+60)break;paint(ms,r,1,2,"soft");paint(ms,r,8,8,"note");}
      for(let k=r;k<r+3;k++)if(/^\s*total/i.test(ms.text("A"+k)))paint(ms,k,1,mc,"total");}
    // onglet Notes
    const head=ns.find(t=>/^\s*nom\s*&\s*pr/i.test(t)&&!/intervenant/i.test(t));
    if(head){const lastTxt=r=>Math.max(1,...ns.rowCells(r).filter(x=>x.text.trim()&&!kid(x.c,"f")).map(x=>x.col));const hasSub=ns.rowCells(head.r+1).some(x=>/^\s*\/\s*\d+/.test(x.text));
      const mc=Math.max(lastTxt(head.r),hasSub?lastTxt(head.r+1):1,3);paint(ns,head.r,1,mc,"header");if(hasSub)paint(ns,head.r+1,1,mc,"header");}
    const moy=ns.find(t=>/^\s*moyenne/i.test(t));if(moy)paint(ns,moy.r,1,ns.maxCol(),"total");
    for(const re of [/^\s*date\s*$/i,/^\s*promotion/i,/situation/i,/^\s*capacit/i,/intervenant/i]){const x=ns.find(t=>re.test(t));if(x)paint(ns,x.r,x.c,x.c,"label");}
  }
  b.zip.file(path,ser(d));
}


/* ---------- Résultats de la promotion : modèles d'Alban (tableau final, PEP, relevés de notes) ---------- */
const RES_CODES=["C1.1","C1.2","C1.3","C2.1","C2.2","C2.3","C2.4","C3.1","C3.2","C3.3","C4.1","C4.2","C4.3","C5.1","C5.2","C5.3","C6.1","C6.2","C6.3","C7.1","C7.2","C7.3","C8.1","C8.2","C8.3"];
const frNum=n=>(Math.round(n*100)/100).toFixed(2).replace(".",",");
function semNum(s){const m=/([1-4])/.exec(String(s||""));return m?+m[1]:0;}
/* Peintre de styles réutilisable (mêmes rôles que les grilles : header, soft, total, note, label) */
function makePainter(d,T){
  const cellXfs=all(d,"cellXfs")[0],fontsEl=all(d,"fonts")[0],fillsEl=all(d,"fills")[0];
  for(const fn of kids(fontsEl,"font")){const nm=kid(fn,"name");if(nm)nm.setAttribute("val",T.font);if(T.size){const s=kid(fn,"sz");if(s&&+s.getAttribute("val")<T.size)s.setAttribute("val",T.size);}}
  const fillCache={},fontCache={},xfCache={};
  const addFill=rgb=>{if(fillCache[rgb]!=null)return fillCache[rgb];const fl=d.createElementNS(NS,"fill"),pf=d.createElementNS(NS,"patternFill");pf.setAttribute("patternType","solid");
    const fg=d.createElementNS(NS,"fgColor");fg.setAttribute("rgb",rgb);const bg=d.createElementNS(NS,"bgColor");bg.setAttribute("indexed","64");pf.appendChild(fg);pf.appendChild(bg);fl.appendChild(pf);fillsEl.appendChild(fl);
    fillsEl.setAttribute("count",kids(fillsEl,"fill").length);return fillCache[rgb]=kids(fillsEl,"fill").length-1;};
  const addFont=(base,sp)=>{const key=base+"|"+JSON.stringify(sp);if(fontCache[key]!=null)return fontCache[key];const fn=kids(fontsEl,"font")[base].cloneNode(true);
    if(sp.b&&!kid(fn,"b")){const e=d.createElementNS(NS,"b");e.setAttribute("val","true");fn.insertBefore(e,fn.firstChild);}
    if(sp.color){let c=kid(fn,"color");if(!c){c=d.createElementNS(NS,"color");const sz=kid(fn,"sz");fn.insertBefore(c,sz?sz.nextSibling:fn.firstChild);}for(const a of ["theme","indexed","tint","auto"])c.removeAttribute(a);c.setAttribute("rgb",sp.color);}
    fontsEl.appendChild(fn);fontsEl.setAttribute("count",kids(fontsEl,"font").length);return fontCache[key]=kids(fontsEl,"font").length-1;};
  const variant=(s,role)=>{const sp=T[role];if(!sp)return s;const key=s+"|"+role;if(xfCache[key]!=null)return xfCache[key];
    const xf=kids(cellXfs,"xf")[s]||kids(cellXfs,"xf")[0];const n=xf.cloneNode(true);
    if(sp.fill){n.setAttribute("fillId",addFill(sp.fill));n.setAttribute("applyFill","true");}
    if(sp.color||sp.b){n.setAttribute("fontId",addFont(+xf.getAttribute("fontId")||0,{b:sp.b,color:sp.color}));n.setAttribute("applyFont","true");}
    cellXfs.appendChild(n);cellXfs.setAttribute("count",kids(cellXfs,"xf").length);return xfCache[key]=kids(cellXfs,"xf").length-1;};
  return {paint(sh,r1,r2,c1,c2,role){for(let r=r1;r<=r2;r++){const row=sh.row(r,false);if(!row)continue;
    for(const c of kids(row,"c")){const col=splitRef(c.getAttribute("r")).c;if(col<c1||col>c2)continue;c.setAttribute("s",variant(+c.getAttribute("s")||0,role));}}}};
}
async function withStyles(b,fn){const path=resolve(b.wbPath,(b.wbRels.find(r=>/styles$/.test(r.type))||{}).target||"styles.xml");
  const d=parse(await b.zip.file(path).async("string"));fn(d);b.zip.file(path,ser(d));}
/* Décale les références de ligne d'une formule (références relatives uniquement) */
function shiftFormula(f,from,to){return f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])/g,(m,d1,c,d2,n)=>(!d2&&+n===from)?d1+c+d2+to:m);}
function setRowNum(row,r){row.setAttribute("r",r);for(const c of kids(row,"c")){const p=splitRef(c.getAttribute("r"));c.setAttribute("r",colStr(p.c)+r);}}

/* Tableau final (« Suivi général des notes ECCF ») */
async function resFinale(buf,o){
  const b=await openBook(buf);const {zip}=b;await fixDrawings(zip);
  const s=b.sheets[0];const sh=new Sheet(parse(await zip.file(s.path).async("string")),b.sst);
  const head=sh.rows().map(r=>+r.getAttribute("r")).find(r=>sh.rowCells(r).filter(x=>/^\s*C\d\.\d/.test(x.text)).length>=20);
  if(!head)throw new Error("ligne des capacités introuvable dans le modèle du tableau final");
  for(const x of sh.rowCells(head)){const m=/^\s*(C\d\.\d)/.exec(x.text);if(!m)continue;const code=m[1];
    if(o.titres&&o.titres[code])sh.setStr(x.ref,code+" "+o.titres[code]);
    const sem=o.semestres&&o.semestres[code];if(sem)sh.setStr(colStr(x.col)+(head-1),"Semestre "+sem);}
  const P=head+2;/* ligne prototype apprenant */const proto=sh.row(P,false);if(!proto)throw new Error("ligne apprenant du modèle introuvable");
  const list=(o.apprenants&&o.apprenants.length)?o.apprenants:[{nom:"",prenom:""}];const N=list.length;
  const after=sh.rows().filter(r=>+r.getAttribute("r")>P);const oldFirst=after.length?+after[0].getAttribute("r"):P+1;
  const delta=(P+N)-oldFirst;
  for(const row of after.slice().reverse()){const r=+row.getAttribute("r");setRowNum(row,r+delta);
    for(const f of all(row,"f")){f.textContent=f.textContent.replace(/(\$?[A-Z]{1,3})(\d+):(\$?[A-Z]{1,3})(\d+)/g,(m,a,ra,c,rb)=>+ra===P?a+ra+":"+c+(P+N-1):m).replace(/\/\s*18\b/,"/"+N);}}
  let prev=proto;for(let i=1;i<N;i++){const row=proto.cloneNode(true);setRowNum(row,P+i);for(const f of all(row,"f"))f.textContent=shiftFormula(f.textContent,P,P+i);prev.parentNode.insertBefore(row,prev.nextSibling);prev=row;}
  list.forEach((a,i)=>{if(a.nom)sh.setStr("A"+(P+i),a.nom);if(a.prenom)sh.setStr("B"+(P+i),a.prenom);});
  if(o.notes){const capCol={};for(const x of sh.rowCells(head)){const m=/^\s*(C\d\.\d)/.exec(x.text);if(m)capCol[m[1]]=x.col;}
    o.notes.forEach((nt,i)=>{if(!nt)return;for(const [code,v] of Object.entries(nt))if(capCol[code]&&v!=null&&v!=="")sh.setNum(colStr(capCol[code])+(P+i),+v);});}
  const mc=all(sh.doc,"mergeCells")[0];if(mc)for(const m of kids(mc,"mergeCell")){const ref=m.getAttribute("ref");
    m.setAttribute("ref",ref.replace(/([A-Z]+)(\d+)/g,(x,c,n)=>+n>P?c+(+n+delta):x));}
  const dim=all(sh.doc,"dimension")[0];if(dim)dim.parentNode.removeChild(dim);
  /* titre (zone de texte) */
  for(const p of Object.keys(zip.files).filter(p=>/^xl\/drawings\/drawing\d+\.xml$/.test(p))){let x=await zip.file(p).async("string");
    if(o.promo)x=x.replace(/(BTS AP\s*)\d{4}\s*-\s*\d{4}/g,"$1"+xmlEsc(o.promo));
    if(o.periode)x=x.replace(/Fin de\s+[Ss]emestre\s*\d/g,"Fin de semestre "+o.periode);
    zip.file(p,x);}
  const T=THEMES[o.theme];
  if(T){const last=P+N-1,avg=P+N+1,mc2=sh.maxCol();const modCols=[];for(const x of sh.rowCells(head-1))if(/^\s*M\d\s*$/.test(x.text))modCols.push(x.col);
    await withStyles(b,d=>{const pt=makePainter(d,T);pt.paint(sh,2,head+1,1,mc2,"header");
      for(const c of modCols)pt.paint(sh,P,last,c,c,"soft");pt.paint(sh,P,last,mc2-3,mc2,"note");pt.paint(sh,P,last,1,2,"label");pt.paint(sh,avg,avg,1,mc2,"total");});}
  zip.file(s.path,ser(sh.doc));
  return await zip.generateAsync({type:"uint8array",compression:"DEFLATE"});
}

/* Plan d'évaluation prévisionnel */
const PEP_COLS=["K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC"];
const PEP_MONTH=[9,10,11,12,1,2,3,4,5,6,9,10,11,12,1,2,3,4,5];
const PEP_YEAR=[0,0,0,0,1,1,1,1,1,1,1,1,1,1,2,2,2,2,2];
const PEP_SEM=[1,1,1,1,1,2,2,2,2,2,3,3,3,3,3,4,4,4,4];
async function resPEP(buf,o){
  const b=await openBook(buf);const {zip}=b;await fixDrawings(zip);
  const s=b.sheets[0];const sh=new Sheet(parse(await zip.file(s.path).async("string")),b.sst);
  const Y=+((/(\d{4})/.exec(o.promo||"")||[])[1])||new Date().getFullYear();
  const t1=sh.find(t=>/PLAN D.EVALUATION/i.test(t));if(t1)sh.setStr(t1.ref,sh.text(t1.ref).replace(/\d{4}\s*-\s*\d{4}/,Y+" - "+(Y+2)));
  const a1=sh.find(t=>/^\s*Ann[ée]e\s*\d{4}/i.test(t));if(a1)sh.setStr(a1.ref,"Année "+Y+"/"+(Y+1));
  const a2=sh.find((t,c)=>/^\s*Ann[ée]e\s*\d{4}/i.test(t)&&c.getAttribute("r")!==(a1&&a1.ref));if(a2)sh.setStr(a2.ref,"Année "+(Y+1)+"/"+(Y+2));
  const mr=sh.find(t=>/^\s*SEPT?\b/i.test(t));const monthRow=mr?mr.r:4;
  PEP_COLS.forEach((c,i)=>{const t=sh.text(c+monthRow);if(t)sh.setStr(c+monthRow,t.replace(/\d{4}/,String(Y+PEP_YEAR[i])));});
  const rowOf={};for(const row of sh.rows()){const r=+row.getAttribute("r");const m=/^\s*(C\d\.\d)/.exec(sh.text("F"+r));if(m&&!rowOf[m[1]])rowOf[m[1]]=r;}
  const placed=[];
  for(const g of o.grilles||[]){const r=rowOf[g.code];if(!r)continue;const sem=semNum(g.semestre);let idx=-1;
    if(g.date){const [yy,mm]=g.date.split("-").map(Number);idx=PEP_COLS.findIndex((c,i)=>PEP_MONTH[i]===mm&&Y+PEP_YEAR[i]===yy);}
    if(idx<0&&sem)idx=PEP_SEM.lastIndexOf(sem);
    if(idx<0)continue;sh.setStr(PEP_COLS[idx]+r,g.code);placed.push([PEP_COLS[idx],r]);}
  const T=THEMES[o.theme];
  if(T){await withStyles(b,d=>{const pt=makePainter(d,T);pt.paint(sh,2,monthRow,1,colNum("AC"),"header");
      const rows=Object.values(rowOf);const r1=Math.min(...rows),r2=Math.max(...rows);pt.paint(sh,r1,r2,1,2,"soft");pt.paint(sh,r1,r2,6,6,"label");
      for(const [c,r] of placed)pt.paint(sh,r,r,colNum(c),colNum(c),"note");});}
  zip.file(s.path,ser(sh.doc));
  return await zip.generateAsync({type:"uint8array",compression:"DEFLATE"});
}

/* Lecture d'un tableau final rempli (dans le navigateur uniquement) */
async function resLire(buf){
  const b=await openBook(buf);const s=b.sheets[0];const sh=new Sheet(parse(await b.zip.file(s.path).async("string")),b.sst);
  const head=sh.rows().map(r=>+r.getAttribute("r")).find(r=>sh.rowCells(r).filter(x=>/^\s*C\d\.\d/.test(x.text)).length>=20);
  if(!head)throw new Error("ce fichier n'est pas un tableau final des notes (ligne des capacités introuvable)");
  const caps=[];for(const x of sh.rowCells(head)){const m=/^\s*(C\d\.\d)\s*(.*)$/s.exec(x.text);if(!m)continue;
    caps.push({code:m[1],titre:m[2].replace(/\s+/g," ").trim(),col:x.col,sem:semNum(sh.text(colStr(x.col)+(head-1)))});}
  const apprenants=[];const maxR=Math.max(...sh.rows().map(r=>+r.getAttribute("r")));
  for(let r=head+1;r<=maxR;r++){const A=sh.text("A"+r).trim(),B=sh.text("B"+r).trim();if(/^moyenne/i.test(A))break;if(!A&&!B)continue;
    const notes={};for(const c of caps){const cell=sh.cell(colStr(c.col)+r,false);if(!cell||kid(cell,"f"))continue;const t=sh.text(cell).trim().replace(",",".");if(t!==""&&isFinite(+t))notes[c.code]=+t;}
    apprenants.push({nom:A,prenom:B,notes});}
  const titre=await (async()=>{for(const p of Object.keys(b.zip.files).filter(p=>/^xl\/drawings\/drawing\d+\.xml$/.test(p))){const x=await b.zip.file(p).async("string");const m=/BTS AP\s*(\d{4}\s*-\s*\d{4})/.exec(x);if(m)return m[1].replace(/\s/g,"");}return "";})();
  return {caps,apprenants,promo:titre};
}

/* Relevé de notes individuel (docx) */
const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DOCX_THEME={origine:{fill:"FFFFFF",color:"000000",font:null},kerplouz:{fill:"2E6A4B",color:"FFFFFF",font:"Calibri"},
  impression:{fill:"E7E7E7",color:"000000",font:"Arial"},logo:{fill:"D1683F",color:"FFFFFF",font:"Century Gothic"}};
const BLOC_LIB=n=>"Bloc "+n+(n<=3?" – Tronc commun":" – Professionnel");
function resCalc(k,caps,notes){
  const list=RES_CODES.map(code=>caps.find(c=>c.code===code)||{code,titre:"",sem:0}).filter(c=>k===4||(c.sem&&c.sem<=k));
  const vals=list.filter(c=>notes[c.code]!=null).map(c=>notes[c.code]);
  const blocs={};for(const c of list){const bn=+c.code[1];(blocs[bn]=blocs[bn]||[]);if(notes[c.code]!=null)blocs[bn].push(notes[c.code]);}
  const bMoy={};for(const [bn,v] of Object.entries(blocs))if(v.length)bMoy[bn]=v.reduce((a,x)=>a+x,0)/v.length;
  let moy=null;if(k===4){const v=Object.values(bMoy);if(v.length)moy=v.reduce((a,x)=>a+x,0)/v.length;}else if(vals.length)moy=vals.reduce((a,x)=>a+x,0)/vals.length;
  const sum=vals.reduce((a,x)=>a+x,0),rest=25-vals.length;const mini=rest>0?Math.max(0,(250-sum)/rest):0;
  return {list,moy,bMoy,mini,nb:vals.length};}
async function resReleve(buf,o){
  const zip=await root.JSZip.loadAsync(buf);const th=DOCX_THEME[o.theme]||DOCX_THEME.origine;
  let x=await zip.file("word/document.xml").async("string");
  const d=parse(x);const k=o.k;const calc=resCalc(k,o.caps,o.notes);
  const trs=all(d,"tr");const proto=trs.find(tr=>/«CODE»/.test(tr.textContent));
  const setT=(el,from,to)=>{for(const t of all(el,"t"))if(t.textContent.includes(from))t.textContent=t.textContent.split(from).join(to);};
  const vmerge=(tc,restart)=>{const pr=kid(tc,"tcPr");const v=d.createElementNS(W,"w:vMerge");if(restart)v.setAttributeNS(W,"w:val","restart");
    const after=kid(pr,"tcW");pr.insertBefore(v,after?after.nextSibling:pr.firstChild);};
  let lastBloc=null;
  for(const c of calc.list){const tr=proto.cloneNode(true);const note=o.notes[c.code];const bn=+c.code[1];
    setT(tr,"«CODE»",c.code);setT(tr,"«INTITULE»",c.titre||"");setT(tr,"«NOTE»",note!=null?frNum(note):"—");
    if(k===4){const tcs=kids(tr,"tc");const first=bn!==lastBloc;
      setT(tr,"«BLOC»",first?BLOC_LIB(bn):"");setT(tr,"«MOYBLOC»",first&&calc.bMoy[bn]!=null?frNum(calc.bMoy[bn]):"");
      vmerge(tcs[0],first);vmerge(tcs[4],first);lastBloc=bn;}
    proto.parentNode.insertBefore(tr,proto);}
  proto.parentNode.removeChild(proto);
  x=ser(d);
  const Y=+((/(\d{4})/.exec(o.promo||"")||[])[1])||0;
  const rep={"«APPRENANT»":o.apprenant,"«PROMO»":o.promo||"","«ANNEE»":Y?Y+"–"+(Y+1):"","«ANNEE2»":Y?(Y+1)+"-"+(Y+2):"","«DATE»":o.date||"",
    "«SIGNATAIRE»":o.signataire||"","«MOY»":calc.moy!=null?frNum(calc.moy):"—","«MINI»":frNum(calc.mini),"«PUBLICATION»":o.publication?" le "+o.publication:""};
  for(const [a,v] of Object.entries(rep))x=x.split(a).join(xmlEsc(v));
  x=x.split("HFILL").join(th.fill).split("HCOLOR").join(th.color);
  if(th.font){x=x.replace(/(w:(?:ascii|hAnsi|cs|eastAsia)=")Arial(")/g,"$1"+th.font+"$2");
    let st=await zip.file("word/styles.xml").async("string");st=st.replace(/(w:(?:ascii|hAnsi|cs|eastAsia)=")Arial(")/g,"$1"+th.font+"$2");zip.file("word/styles.xml",st);}
  zip.file("word/document.xml",x);
  return await zip.generateAsync({type:"uint8array",compression:"DEFLATE"});
}
/* Courrier générique (convocation, résultats, alerte) : champs «X», lignes de tableau «CODE», puces «PUCE» */
async function resDoc(buf,o){
  const zip=await root.JSZip.loadAsync(buf);const th=DOCX_THEME[o.theme]||DOCX_THEME.origine;
  const d=parse(await zip.file("word/document.xml").async("string"));
  const setT=(el,from,to)=>{for(const t of all(el,"t"))if(t.textContent.includes(from))t.textContent=t.textContent.split(from).join(to);};
  const proto=all(d,"tr").find(tr=>/«CODE»/.test(tr.textContent));
  if(proto){const L=(o.lignes&&o.lignes.length)?o.lignes:[{CODE:"—",INTITULE:"Aucune",NOTE:""}];
    for(const l of L){const tr=proto.cloneNode(true);for(const k of new Set(["CODE","INTITULE","NOTE",...Object.keys(l)]))setT(tr,"«"+k+"»",l[k]==null?"":String(l[k]));proto.parentNode.insertBefore(tr,proto);}
    proto.parentNode.removeChild(proto);}
  const pp=all(d,"p").find(p=>/«PUCE»/.test(p.textContent));
  if(pp){for(const t of (o.puces||[])){const p=pp.cloneNode(true);setT(p,"«PUCE»",t);pp.parentNode.insertBefore(p,pp);}pp.parentNode.removeChild(pp);}
  let x=ser(d);
  for(const [k,v] of Object.entries(o.champs||{}))x=x.split("«"+k+"»").join(xmlEsc(v==null?"":String(v)));
  x=x.replace(/«[A-Z0-9]+»/g,"");
  x=x.split("HFILL").join(th.fill==="FFFFFF"?"EEF3EF":th.fill).split("HCOLOR").join(th.color);
  if(th.font){x=x.replace(/(w:(?:ascii|hAnsi|cs|eastAsia)=")Arial(")/g,"$1"+th.font+"$2");
    let st=await zip.file("word/styles.xml").async("string");st=st.replace(/(w:(?:ascii|hAnsi|cs|eastAsia)=")Arial(")/g,"$1"+th.font+"$2");zip.file("word/styles.xml",st);}
  zip.file("word/document.xml",x);
  return await zip.generateAsync({type:"uint8array",compression:"DEFLATE"});
}
root.ResultatsEngine={finale:resFinale,pep:resPEP,lire:resLire,releve:resReleve,doc:resDoc,calc:resCalc,CODES:RES_CODES,frNum};

root.GrillesEngine={inspect,fill,titleCase,THEMES,pwHash,gridStructure:(x)=>x};
})(typeof window!=="undefined"?window:globalThis);
