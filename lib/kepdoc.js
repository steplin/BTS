/* KepDoc \u2014 une m\u00eame grille produit un classeur Excel (.xlsx) et un PDF (sans impression HTML).
   Feuille : {nom, papier:"A3"|"A4", paysage, marges:{g,d,h,b} (pouces), cols:[largeurs Excel], lignes:[{h, cells:[cellule|null], entete}],
              teteXL:{g,c,d}, piedXL:{g,d}, hautPDF:[lignes], piedPDF:"texte gauche", repeter:n (lignes d'en-t\u00eate r\u00e9p\u00e9t\u00e9es)}
   Cellule : {v, b, i, sz, al:"l"|"c"|"r", va:"t"|"c", fond:"3C7D15", coul:"FFFFFF", bord:true|false, rs, cs}
   D\u00e9pendances charg\u00e9es \u00e0 la demande : /lib/jszip.min.js, /lib/pdf-lib.min.js */
(function(){
"use strict";
const charger=src=>new Promise((ok,ko)=>{if(document.querySelector(`script[data-kep="${src}"]`)){const t=()=>((src.includes("jszip")&&window.JSZip)||(src.includes("pdf-lib")&&window.PDFLib))?ok():setTimeout(t,50);return t();}
  const s=document.createElement("script");s.src=src;s.dataset.kep=src;s.onload=()=>ok();s.onerror=()=>ko(new Error("Chargement impossible : "+src));document.head.appendChild(s);});
const zipLib=()=>window.JSZip?Promise.resolve():charger("/lib/jszip.min.js");
const pdfLib=()=>window.PDFLib?Promise.resolve():charger("/lib/pdf-lib.min.js");
const PT_COL=5.25;                          // 1 unit\u00e9 de largeur Excel \u2248 7 px \u2248 5,25 pt
const txt=v=>v==null?"":String(v);
const xesc=s=>txt(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,"");

/* ---------- Hauteur de ligne estim\u00e9e (Excel ne recalcule pas les lignes fusionn\u00e9es) ---------- */
function lignesEstimees(v,largPt,sz){const cw=sz*0.5;let n=0;for(const p of txt(v).split("\n")){n+=Math.max(1,Math.ceil((p.length*cw)/Math.max(10,largPt-4)));}return n;}
function hauteurs(F){const H=F.lignes.map(l=>l.h||0),larg=F.cols.map(w=>w*PT_COL);
  F.lignes.forEach((l,r)=>{let c=0;l.cells.forEach((x,ci)=>{if(x&&!(x.rs>1)){const w=larg.slice(ci,ci+(x.cs||1)).reduce((a,b)=>a+b,0),sz=x.sz||10;
      H[r]=Math.max(H[r],lignesEstimees(x.v,w,sz)*sz*1.22+5);}});});
  F.lignes.forEach((l,r)=>l.cells.forEach((x,ci)=>{if(x&&x.rs>1){const w=larg.slice(ci,ci+(x.cs||1)).reduce((a,b)=>a+b,0),sz=x.sz||10,need=lignesEstimees(x.v,w,sz)*sz*1.22+5;
      const tot=H.slice(r,r+x.rs).reduce((a,b)=>a+b,0);if(need>tot)H[r+x.rs-1]+=need-tot;}}));
  return H.map(h=>Math.max(13,Math.min(409,Math.round(h*10)/10)));}

/* ---------- Excel (.xlsx) ---------- */
async function xlsx(feuilles,props={}){await zipLib();const Z=new JSZip();
  const fonts=['<font><sz val="10"/><name val="Arial"/><family val="2"/></font>'],fills=['<fill><patternFill patternType="none"/></fill>','<fill><patternFill patternType="gray125"/></fill>'],
    borders=['<border><left/><right/><top/><bottom/><diagonal/></border>','<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>'],
    xfs=['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'],idx=new Map(),fIdx=new Map(),flIdx=new Map();
  const style=x=>{const sz=x.sz||10,fk=[sz,!!x.b,!!x.i,x.coul||""].join("|");
    if(!fIdx.has(fk)){fIdx.set(fk,fonts.length);fonts.push(`<font>${x.b?"<b/>":""}${x.i?"<i/>":""}<sz val="${sz}"/>${x.coul?`<color rgb="FF${x.coul}"/>`:""}<name val="Arial"/><family val="2"/></font>`);}
    let fl=0;if(x.fond){if(!flIdx.has(x.fond)){flIdx.set(x.fond,fills.length);fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="FF${x.fond}"/><bgColor indexed="64"/></patternFill></fill>`);}fl=flIdx.get(x.fond);}
    const bd=x.bord===false?0:1,h={l:"left",c:"center",r:"right"}[x.al||"l"],v=x.va==="t"?"top":"center",k=[fIdx.get(fk),fl,bd,h,v].join("|");
    if(!idx.has(k)){idx.set(k,xfs.length);xfs.push(`<xf numFmtId="0" fontId="${fIdx.get(fk)}" fillId="${fl}" borderId="${bd}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="${h}" vertical="${v}" wrapText="1"/></xf>`);}
    return idx.get(k);};
  const col=n=>{let s="";n++;while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;};
  const noms=[],defs=[];
  feuilles.forEach((F,si)=>{const H=hauteurs(F),merges=[],nc=F.cols.length;const occ=new Set();
    let rows="";F.lignes.forEach((l,r)=>{let cells="";
      for(let c=0;c<nc;c++){const x=l.cells[c];const ref=col(c)+(r+1);
        if(x){const s=style(x);cells+=`<c r="${ref}" s="${s}"${txt(x.v)!==""?` t="inlineStr"><is><t xml:space="preserve">${xesc(x.v)}</t></is></c>`:"/>"}`;
          if((x.rs||1)>1||(x.cs||1)>1){merges.push(ref+":"+col(c+(x.cs||1)-1)+(r+(x.rs||1)));
            for(let rr=r;rr<r+(x.rs||1);rr++)for(let cc=c;cc<c+(x.cs||1);cc++)if(rr!==r||cc!==c)occ.add(rr+"|"+cc+"|"+s);}}
        else{const o=[...occ].find(k=>k.startsWith(r+"|"+c+"|"));if(o)cells+=`<c r="${ref}" s="${o.split("|")[2]}"/>`;}}
      rows+=`<row r="${r+1}" ht="${H[r]}" customHeight="1">${cells}</row>`;});
    const m=F.marges||{g:.3,d:.3,h:.9,b:.7},paper=F.papier==="A3"?8:9;
    const hf=(F.teteXL||F.piedXL)?`<headerFooter><oddHeader>${xesc((F.teteXL?`&L${F.teteXL.g||""}&C${F.teteXL.c||""}&R${F.teteXL.d||""}`:""))}</oddHeader><oddFooter>${xesc(F.piedXL?`&L${F.piedXL.g||""}&R${F.piedXL.d||"Page &P sur &N"}`:"")}</oddFooter></headerFooter>`:"";
    Z.file(`xl/worksheets/sheet${si+1}.xml`,`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView workbookViewId="0"${si===0?' tabSelected="1"':""}/></sheetViews><sheetFormatPr defaultRowHeight="13"/>
<cols>${F.cols.map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join("")}</cols><sheetData>${rows}</sheetData>${merges.length?`<mergeCells count="${merges.length}">${merges.map(x=>`<mergeCell ref="${x}"/>`).join("")}</mergeCells>`:""}
<printOptions horizontalCentered="1"/><pageMargins left="${m.g}" right="${m.d}" top="${m.h}" bottom="${m.b}" header="0.3" footer="0.3"/><pageSetup paperSize="${paper}" orientation="${F.paysage?"landscape":"portrait"}" fitToWidth="1" fitToHeight="0"/>${hf}</worksheet>`);
    const nom=(F.nom||"Feuille "+(si+1)).replace(/[\\\/\?\*\[\]:]/g," ").slice(0,31);noms.push(nom);
    if(F.repeter)defs.push(`<definedName name="_xlnm.Print_Titles" localSheetId="${si}">'${nom.replace(/'/g,"''")}'!$${F.repeterDe||1}:$${(F.repeterDe||1)+F.repeter-1}</definedName>`);});
  Z.file("[Content_Types].xml",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${feuilles.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`);
  Z.file("_rels/.rels",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`);
  Z.file("docProps/core.xml",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xesc(props.titre||"")}</dc:title><dc:creator>Kerplouz LaSalle</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().slice(0,19)}Z</dcterms:created></cp:coreProperties>`);
  Z.file("xl/workbook.xml",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${noms.map((n,i)=>`<sheet name="${xesc(n)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join("")}</sheets>${defs.length?`<definedNames>${defs.join("")}</definedNames>`:""}</workbook>`);
  Z.file("xl/_rels/workbook.xml.rels",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${noms.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join("")}<Relationship Id="rId${noms.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  Z.file("xl/styles.xml",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="${fonts.length}">${fonts.join("")}</fonts><fills count="${fills.length}">${fills.join("")}</fills><borders count="${borders.length}">${borders.join("")}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
  return Z.generateAsync({type:"uint8array",mimeType:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});}

/* ---------- Polices DejaVu (sous-ensembles /fonts/pdf-*.ttf) : lecteur TrueType minimal pour pdf-lib ---------- */
const MiniFontkit={create(buf){const u8=buf instanceof Uint8Array?buf:new Uint8Array(buf),dv=new DataView(u8.buffer,u8.byteOffset,u8.byteLength);
  const n=dv.getUint16(4),T={};for(let i=0;i<n;i++){const o=12+i*16,tag=String.fromCharCode(u8[o],u8[o+1],u8[o+2],u8[o+3]);T[tag]={off:dv.getUint32(o+8),len:dv.getUint32(o+12)};}
  const head=T.head.off,hhea=T.hhea.off,upm=dv.getUint16(head+18),bbox={minX:dv.getInt16(head+36),minY:dv.getInt16(head+38),maxX:dv.getInt16(head+40),maxY:dv.getInt16(head+42)},macStyle=dv.getUint16(head+44);
  const ascent=dv.getInt16(hhea+4),descent=dv.getInt16(hhea+6),nhm=dv.getUint16(hhea+34),numG=dv.getUint16(T.maxp.off+4);
  const adv=new Array(numG);let last=0;for(let g=0;g<numG;g++){if(g<nhm)last=dv.getUint16(T.hmtx.off+g*4);adv[g]=last;}
  const post=T.post.off,italicAngle=dv.getInt32(post+4)/65536,isFixedPitch=dv.getUint32(post+12)!==0;
  let capHeight=0,xHeight=0,sFamilyClass=0;if(T["OS/2"]){const o=T["OS/2"].off,v=dv.getUint16(o);sFamilyClass=dv.getInt16(o+30);if(v>=2){xHeight=dv.getInt16(o+86);capHeight=dv.getInt16(o+88);}}
  // cmap format 4 (Windows Unicode BMP)
  const map=new Map(),cm=T.cmap.off,nt=dv.getUint16(cm+2);let sub=-1;
  for(let i=0;i<nt;i++){const p=dv.getUint16(cm+4+i*8),e=dv.getUint16(cm+6+i*8),o=dv.getUint32(cm+8+i*8);if(dv.getUint16(cm+o)===4&&(p===3&&e===1||p===0))sub=cm+o;}
  if(sub>=0){const segX2=dv.getUint16(sub+6),ends=sub+14,starts=ends+segX2+2,deltas=starts+segX2,ranges=deltas+segX2;
    for(let s=0;s<segX2/2;s++){const end=dv.getUint16(ends+s*2),start=dv.getUint16(starts+s*2),delta=dv.getInt16(deltas+s*2),ro=dv.getUint16(ranges+s*2);
      for(let c=start;c<=end&&c!==0xFFFF;c++){let g;if(ro===0)g=(c+delta)&0xFFFF;else{const gi=ranges+s*2+ro+(c-start)*2;g=dv.getUint16(gi);if(g)g=(g+delta)&0xFFFF;}if(g)map.set(c,g);}}}
  let psName="DejaVu";if(T.name){const o=T.name.off,cnt=dv.getUint16(o+2),so=o+dv.getUint16(o+4);
    for(let i=0;i<cnt;i++){const r=o+6+i*12,pid=dv.getUint16(r),nid=dv.getUint16(r+6),l=dv.getUint16(r+8),off=dv.getUint16(r+10);
      if(nid===6){let s="";for(let k=0;k<l;k+=(pid===1?1:2))s+=String.fromCharCode(pid===1?u8[so+off+k]:dv.getUint16(so+off+k));if(s){psName=s;break;}}}}
  const cache=new Map();const glyph=(cp)=>{const id=map.get(cp)||0,k=id+"|"+cp;if(!cache.has(k))cache.set(k,{id,codePoints:[cp],advanceWidth:adv[id]||0,name:""});return cache.get(k);};
  return {unitsPerEm:upm,bbox,ascent,descent,italicAngle,capHeight,xHeight,postscriptName:psName,cff:false,post:{isFixedPitch},head:{macStyle:{italic:!!(macStyle&2)}},"OS/2":{sFamilyClass},
    characterSet:[...map.keys()].sort((a,b)=>a-b),glyphForCodePoint:glyph,hasGlyphForCodePoint:cp=>map.has(cp),
    layout(t){const g=[];for(const ch of String(t))g.push(glyph(ch.codePointAt(0)));return {glyphs:g,positions:g.map(x=>({xAdvance:x.advanceWidth}))};}};}};
const POLICES={r:"pdf-DejaVuSansCondensed",b:"pdf-DejaVuSansCondensed-Bold",sr:"pdf-DejaVuSerifCondensed",sb:"pdf-DejaVuSerifCondensed-Bold",si:"pdf-DejaVuSerifCondensed-Italic"};
let policesOctets=null;
async function chargerPolices(){if(policesOctets)return policesOctets;const o={};
  await Promise.all(Object.entries(POLICES).map(async([k,f])=>{const r=await fetch("/fonts/"+f+".ttf");if(!r.ok)throw new Error("Police introuvable : "+f);o[k]=new Uint8Array(await r.arrayBuffer());}));
  return policesOctets=o;}

/* ---------- PDF (pdf-lib) : mise en page des grilles officielles (export EvalUc / mPDF) ---------- */
const PAGES={A3:[841.89,1190.55],A4:[595.28,841.89]};
function rgb(h){const n=parseInt(h,16);return PDFLib.rgb(((n>>16)&255)/255,((n>>8)&255)/255,(n&255)/255);}
async function pdf(feuilles,props={}){await pdfLib();const {PDFDocument,rgb:RGB}=PDFLib;
  const doc=await PDFDocument.create();doc.registerFontkit(MiniFontkit);doc.setTitle(props.titre||"");doc.setCreator("Kerplouz LaSalle");doc.setProducer("keprojects");
  const oct=await chargerPolices(),F={};for(const k of Object.keys(POLICES))F[k]=await doc.embedFont(oct[k],{subset:false});
  const jeu=new Set(MiniFontkit.create(oct.r).characterSet);
  const REMP={"\u2018":"'","\u2019":"\u2019","\u00a0":" ","\u202f":" ","\u2009":" ","\u2011":"-","\u00ad":"","\t":" ","\u2713":"x","\u2714":"x"};
  const propre=s=>{let o="";for(const ch of txt(s).normalize("NFC")){if(ch==="\n"){o+=ch;continue;}const cp=ch.codePointAt(0);if(jeu.has(cp)){o+=ch;continue;}
      if(REMP[ch]!==undefined){o+=REMP[ch];continue;}const base=ch.normalize("NFD")[0];o+=jeu.has(base.codePointAt(0))?base:"?";}return o;};
  const fontDe=x=>x.serif?(x.b?F.sb:x.i?F.si:F.sr):(x.b?F.b:F.r);
  const couper=(s,font,sz,larg)=>{const out=[];for(const par of propre(s).split("\n")){if(!par.trim()){out.push("");continue;}let cur="";
      for(const mot of par.split(/ +/)){let m=mot;const essai=cur?cur+" "+m:m;if(font.widthOfTextAtSize(essai,sz)<=larg){cur=essai;continue;}
        if(cur)out.push(cur);cur="";while(font.widthOfTextAtSize(m,sz)>larg&&m.length>1){let k=m.length-1;while(k>1&&font.widthOfTextAtSize(m.slice(0,k),sz)>larg)k--;out.push(m.slice(0,k));m=m.slice(k);}cur=m;}
      out.push(cur);}while(out.length>1&&out[out.length-1]==="")out.pop();return out;};
  const pages=[];
  for(const Fe of feuilles){const P=Fe.pdf||{},[W0,H0]=PAGES[Fe.papier||"A4"],[PW,PH]=Fe.paysage?[H0,W0]:[W0,H0];
    const mg=P.marge||22,md=P.marge||22,largUtile=PW-mg-md,brut=Fe.cols.reduce((a,w)=>a+w,0);
    const unite=P.unite||Math.min(4.79,largUtile/brut),pol=P.police||unite/4.79,cw=Fe.cols.map(w=>w*unite),PAD=P.pad!=null?P.pad:1.6;
    const X=[mg];cw.forEach(w=>X.push(X[X.length-1]+w));
    const L=Fe.lignes.filter(l=>l.pdf!==false).map(l=>({...l,cells:l.cells.map(x=>x?{...x,sz:x.psz||(x.sz||10)*pol}:x)}));
    const prep=x=>{if(!x)return;const f=fontDe(x),w=cw.slice(x.c0,x.c0+(x.cs||1)).reduce((a,b)=>a+b,0)-2*PAD;x.f=f;x.L=couper(x.v,f,x.sz,Math.max(8,w));x.lh=x.sz*1.22;x.need=(txt(x.v)===""?0:x.L.length*x.lh)+2*PAD+1;};
    L.forEach(l=>l.cells.forEach((x,ci)=>{if(x){x.c0=ci;prep(x);}}));
    const H=L.map(l=>Math.max(l.h?l.h*pol*0.92:0,...l.cells.filter(x=>x&&!(x.rs>1)).map(x=>x.need),4));
    L.forEach((l,r)=>l.cells.forEach(x=>{if(x&&x.rs>1){const t=H.slice(r,r+x.rs).reduce((a,b)=>a+b,0);if(x.need>t){const extra=(x.need-t)/x.rs;for(let q=r;q<r+x.rs;q++)H[q]+=extra;}}}));
    const fin=L.map((_,r)=>r);L.forEach((l,r)=>l.cells.forEach(x=>{if(x&&x.rs>1)for(let q=r;q<r+x.rs;q++)fin[q]=Math.max(fin[q],r+x.rs-1);}));
    const blocs=[];for(let r=0;r<L.length;){let e=fin[r];for(let q=r;q<=e;q++)e=Math.max(e,fin[q]);blocs.push([r,e]);r=e+1;}
    const tete=P.repeter?Array.from({length:P.repeter},(_,i)=>(P.repeterDe||1)-1+i):[];
    const E=P.entete||null,hautBas=P.bas||50;
    let page,y;
    const dessinerCellule=(x,x0,y0,w,h,lignesTxt)=>{if(x.fond)page.drawRectangle({x:x0,y:y0-h,width:w,height:h,color:rgb(x.fond)});
      if(x.bord!==false)page.drawRectangle({x:x0,y:y0-h,width:w,height:h,borderColor:RGB(0,0,0),borderWidth:0.5});
      const T=lignesTxt||x.L;if(!T||!T.length||(T.length===1&&T[0]===""))return;const th=T.length*x.lh,col=x.coul?rgb(x.coul):RGB(0,0,0);
      let ty=x.va==="t"?y0-PAD-x.sz*0.95:y0-(h-th)/2-x.sz*0.95;
      for(const t of T){if(ty<y0-h-1)break;const tw=x.f.widthOfTextAtSize(t,x.sz);const tx=x.al==="c"?x0+(w-tw)/2:x.al==="r"?x0+w-PAD-tw:x0+PAD;
        if(t)page.drawText(t,{x:tx,y:ty,size:x.sz,font:x.f,color:col});ty-=x.lh;}};
    const dessinerLigne=(l,hl,y0)=>{l.cells.forEach((x,ci)=>{if(!x||x.rs>1)return;const w=cw.slice(ci,ci+(x.cs||1)).reduce((a,b)=>a+b,0);dessinerCellule(x,X[ci],y0,w,hl);});};
    const dessinerPlage=(r0,r1)=>{const Y0={};for(let r=r0;r<=r1;r++){Y0[r]=y;dessinerLigne(L[r],H[r],y);y-=H[r];}
      for(let r=r0;r<=r1;r++)L[r].cells.forEach((x,ci)=>{if(x&&x.rs>1){const h=H.slice(r,Math.min(r1,r+x.rs-1)+1).reduce((s,v)=>s+v,0),w=cw.slice(ci,ci+(x.cs||1)).reduce((s,v)=>s+v,0);dessinerCellule(x,X[ci],Y0[r],w,h);}});};
    let premiere=true,prochaine=0;
    const nouvellePage=()=>{page=doc.addPage([PW,PH]);pages.push({page,Fe,groupe:P.groupe||Fe,mg,md,PW,PH});y=PH-(P.haut||28);
      if(E){// ligne italique (gauche / centre / droite), puis cadre du candidat : r\u00e9p\u00e9t\u00e9s sur chaque page
        const lw=PW-mg-md,sz=9,bl=y-sz*0.93,cols=[[E.g,mg,lw*0.4,"l"],[E.c,mg+lw*0.4,lw*0.3,"c"],[E.d,mg+lw*0.7,lw*0.3,"r"]];
        for(const [t,x0,w,al] of cols){if(!t)continue;const s=propre(t),tw=F.si.widthOfTextAtSize(s,sz);page.drawText(s,{x:al==="c"?x0+(w-tw)/2:al==="r"?x0+w-tw:x0,y:bl,size:sz,font:F.si});}
        if(E.boite&&E.boite.length){const top=PH-46,bh=21;let x0=mg;
          for(const b of E.boite){const w=lw*b.f;page.drawRectangle({x:x0,y:top-bh,width:w,height:bh,color:rgb("F9F9F9"),borderColor:RGB(0,0,0),borderWidth:0.5});
            const s=propre(b.v),bsz=10.5,tw=F.sb.widthOfTextAtSize(s,bsz);page.drawText(s,{x:x0+(w-tw)/2,y:top-bh/2-bsz*0.35,size:bsz,font:F.sb});x0+=w;}
          y=top-bh-(P.apresBoite!=null?P.apresBoite:5);}
        else y=PH-46;}
      if(!premiere&&tete.length&&(P.finRepeter==null||prochaine<P.finRepeter))dessinerPlage(tete[0],tete[tete.length-1]);premiere=false;};
    const bas=()=>hautBas;
    nouvellePage();
    const hTete=tete.reduce((s,r)=>s+H[r],0);const yHaut=()=>E?(E.boite&&E.boite.length?PH-46-21-5:PH-46):PH-(P.haut||28);
    for(const [a,e] of blocs){prochaine=a;const hb=H.slice(a,e+1).reduce((s,v)=>s+v,0);
      if(L[a].saut&&y<yHaut()-1)nouvellePage();
      const suite=L[a].avecSuivant?H.slice(e+1,e+2).reduce((s,v)=>s+v,0):0;
      if(y-hb-suite<bas()&&y<yHaut()-hTete-2)nouvellePage();
      if(y-hb>=bas()){dessinerPlage(a,e);continue;}
      // bloc plus haut qu'une page : coupure entre deux lignes, le texte des cellules fusionn\u00e9es continue
      const fus=[];for(let r=a;r<=e;r++)L[r].cells.forEach((x,ci)=>{if(x&&x.rs>1)fus.push({x,ci,r0:r,r1:r+x.rs-1,reste:x.L.slice()});});
      let debutSeg=y,rSeg=a;const Y={};
      const fermerSeg=(rLast,yFin)=>{for(const f of fus){if(f.r1<rSeg||f.r0>rLast)continue;const top=f.r0>=rSeg?Y[f.r0]:debutSeg,bottom=f.r1<=rLast?Y[f.r1]-H[f.r1]:yFin;
          const h=top-bottom,w=cw.slice(f.ci,f.ci+(f.x.cs||1)).reduce((s,v)=>s+v,0),nn=Math.max(0,Math.floor((h-2*PAD)/f.x.lh)),part=f.reste.splice(0,f.r1<=rLast?f.reste.length:nn);
          dessinerCellule({...f.x,va:f.r1<=rLast&&f.r0>=rSeg?f.x.va:"t"},X[f.ci],top,w,h,part);}};
      for(let r=a;r<=e;r++){if(y-H[r]<bas()){fermerSeg(r-1,y);prochaine=r;nouvellePage();debutSeg=y;rSeg=r;}
        Y[r]=y;dessinerLigne(L[r],H[r],y);y-=H[r];}
      fermerSeg(e,y);}}
  // pied de page : \u00e9tablissement \u00e0 gauche, n/N \u00e0 droite, num\u00e9rotation par document (candidat)
  const tot=new Map();pages.forEach(p=>tot.set(p.groupe,(tot.get(p.groupe)||0)+1));const vu=new Map();
  for(const p of pages){const n=(vu.get(p.groupe)||0)+1;vu.set(p.groupe,n);const P=p.Fe.pdf||{},sz=10.5,yb=P.piedY||29;
    if(P.pied!==undefined||p.Fe.piedPDF){const s=propre(P.pied!==undefined?P.pied:p.Fe.piedPDF);if(s)p.page.drawText(s,{x:p.mg,y:yb,size:sz,font:F.sr});}
    const t=n+"/"+tot.get(p.groupe);p.page.drawText(t,{x:p.PW-p.md-F.sr.widthOfTextAtSize(t,sz),y:yb,size:sz,font:F.sr});}
  return doc.save();}

/* ---------- Courriers (convocations) : un PDF, une lettre par destinataire, mise en page du mod\u00e8le Word ---------- */
async function courriers(docs,o={}){await pdfLib();const {PDFDocument,StandardFonts,rgb:RGB}=PDFLib;
  const doc=await PDFDocument.create();doc.setTitle(o.titre||"");doc.setCreator("Kerplouz LaSalle");doc.setProducer("keprojects");
  const R=await doc.embedFont(StandardFonts.Helvetica),B=await doc.embedFont(StandardFonts.HelveticaBold);
  let logo=null;if(o.logo){try{const r=await fetch(o.logo);if(r.ok)logo=await doc.embedPng(new Uint8Array(await r.arrayBuffer()));}catch(e){}}
  const jeu=new Set(R.getCharacterSet()),REMP={"\u202f":" ","\u00a0":" ","\u2011":"-","\u2010":"-","\u2212":"-","\u2264":"<=","\u2265":">=","\u2192":"->","\u00b7":"\u00b7"};
  const propre=s=>{let out="";for(const ch of txt(s).normalize("NFC")){if(ch==="\n"){out+=ch;continue;}const cp=ch.codePointAt(0);if(jeu.has(cp)){out+=ch;continue;}
      if(REMP[ch]!==undefined&&[...REMP[ch]].every(c=>jeu.has(c.codePointAt(0)))){out+=REMP[ch];continue;}
      if(ch==="\u00b7"){out+="-";continue;}const b=ch.normalize("NFD")[0];out+=jeu.has(b.codePointAt(0))?b:"?";}return out;};
  const pm=s=>txt(s);
  const couper=(s,f,sz,w)=>{const out=[];for(const par of propre(pm(s)).split("\n")){let cur="";for(const m of par.split(/ +/)){const e=cur?cur+" "+m:m;if(f.widthOfTextAtSize(e,sz)<=w){cur=e;continue;}if(cur)out.push(cur);cur=m;}out.push(cur);}return out;};
  const W=595.28,H=841.89,ML=62.35,MR=62.35,MT=51,MB=51,LW=W-ML-MR,th=o.theme||{fill:"FFFFFF",color:"000000"};
  const hex=h=>{const n=parseInt(h,16);return RGB(((n>>16)&255)/255,((n>>8)&255)/255,(n&255)/255);};
  for(const d of docs){let page=doc.addPage([W,H]),y=H-MT;
    const place=h=>{if(y-h<MB){page=doc.addPage([W,H]);y=H-MT;}};
    const para=(t,{f=R,sz=11,al="l",apres=8,x=ML,w=LW}={})=>{const L=couper(t,f,sz,w),lh=sz*1.15;
      for(const l of L){place(lh);const tw=f.widthOfTextAtSize(l,sz);page.drawText(l,{x:al==="r"?x+w-tw:al==="c"?x+(w-tw)/2:x,y:y-sz*0.9,size:sz,font:f});y-=lh;}y-=apres;};
    // en-t\u00eate : logo \u00e0 gauche, destinataire \u00e0 droite
    const top=y;let yg=top;if(logo){const lw=119,lh=lw*logo.height/logo.width;page.drawImage(logo,{x:ML,y:top-lh,width:lw,height:lh});yg=top-lh-4;}
    if(o.qualiopi){const Q=o.qualiopi;for(const l of Q.lignes||[]){page.drawText(propre(l),{x:ML,y:yg-7.5,size:7.5,font:R,color:RGB(.25,.25,.25)});yg-=9;}
        if(Q.code){page.drawText(propre(Q.code),{x:ML,y:yg-8,size:7.5,font:B});yg-=10;}}
    else if(logo){const s=propre(o.etab||"UFA Kerplouz LaSalle Auray");page.drawText(s,{x:ML+2,y:yg-8,size:8.5,font:R});yg-=10;}
    let yd=top;(d.dest||[]).forEach((l,i)=>{if(!l)return;const f=i===0?B:R,s=propre(l),tw=f.widthOfTextAtSize(s,11);page.drawText(s,{x:W-MR-tw,y:yd-10,size:11,font:f});yd-=12.6;});
    y=Math.min(yg,yd)-14;
    para(d.lieuDate||"",{al:"r",apres:16});
    para(d.objet||"",{f:B,apres:12});
    for(const p of d.corps||[])para(p,{apres:8});
    // tableau
    if(d.tableau){const T=d.tableau,n=T.tete.length,cw=T.largeurs?T.largeurs.map(x=>x*LW):Array(n).fill(LW/n),PAD=3,sz=10,lh=sz*1.15;
      const ligne=(cells,tete)=>{const Ls=cells.map((c,i)=>couper(c,tete||(T.gras||[]).includes(i)?B:R,sz,cw[i]-2*PAD));const h=Math.max(...Ls.map(L=>L.length))*lh+2*PAD+2;place(h);let x=ML;
        cells.forEach((c,i)=>{if(tete&&th.fill&&th.fill!=="FFFFFF")page.drawRectangle({x,y:y-h,width:cw[i],height:h,color:hex(th.fill)});
          page.drawRectangle({x,y:y-h,width:cw[i],height:h,borderColor:RGB(0,0,0),borderWidth:0.6});
          const f=tete||(T.gras||[]).includes(i)?B:R;let ty=y-PAD-sz*0.9-(h-2*PAD-Ls[i].length*lh)/2;
          for(const l of Ls[i]){const tw=f.widthOfTextAtSize(l,sz),al=tete||(T.centre||[]).includes(i)?"c":"l";page.drawText(l,{x:al==="c"?x+(cw[i]-tw)/2:x+PAD,y:ty,size:sz,font:f,color:tete?hex(th.color||"000000"):RGB(0,0,0)});ty-=lh;}x+=cw[i];});y-=h;};
      ligne(T.tete,true);for(const r of T.lignes)ligne(r,false);y-=14;}
    if(d.puces&&d.puces.length){para(d.titrePuces||"",{f:B,apres:4});for(const p of d.puces){const L=couper(p,R,11,LW-18);L.forEach((l,i)=>{place(13);if(i===0)page.drawText("\u2022",{x:ML+4,y:y-9.9,size:11,font:R});page.drawText(l,{x:ML+18,y:y-9.9,size:11,font:R});y-=12.65;});}y-=10;}
    for(const p of d.fin||[])para(p,{apres:8});
    y-=24;para(d.signataire||"",{x:ML+255,w:LW-255,apres:40});
    if(d.copie){place(12);const lab="Copie : ";page.drawText(lab,{x:ML,y:y-8,size:8.5,font:B});const L=couper(d.copie,R,8.5,LW-B.widthOfTextAtSize(lab,8.5));
      L.forEach((l,i)=>{page.drawText(l,{x:ML+B.widthOfTextAtSize(lab,8.5),y:y-8-i*10,size:8.5,font:R});});y-=L.length*10;}}
  return doc.save();}

/* ---------- R\u00e9f\u00e9rences Qualiopi des courriers (mod\u00e8les de l'UFA) ---------- */
const ETAB_LIGNES=["AFG KERPLOUZ LaSalle AURAY","Route du Bono \u2013 BP 40417 \u2013 56404 AURAY CEDEX","N\u00b0 d\u00e9claration d\u2019activit\u00e9 53560930256","SIRET 381 125 731 00012 \u2013 CODE NAF 8532Z"];
function qualiopi(ref){if(!ref||!ref.code)return null;const d=ref.date_ref?ref.date_ref.split("-").reverse().join("/"):"";
  return {lignes:ETAB_LIGNES.slice(),code:[ref.code,d,ref.redacteur?"r\u00e9dacteur "+ref.redacteur:""].filter(Boolean).join(" : ")};}
// Ajoute le bloc \u00e9tablissement et la r\u00e9f\u00e9rence en t\u00eate d'un document Word d\u00e9j\u00e0 produit
async function qualiopiDocx(u8,q){if(!q)return u8;await zipLib();const Z=await JSZip.loadAsync(u8),f=Z.file("word/document.xml");if(!f)return u8;let x=await f.async("string");
  const par=(t,b)=>`<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>${b?"<w:b/>":""}<w:color w:val="${b?"000000":"404040"}"/><w:sz w:val="15"/></w:rPr><w:t xml:space="preserve">${xesc(t)}</w:t></w:r></w:p>`;
  const bloc=q.lignes.map(l=>par(l,false)).join("")+par(q.code,true)+`<w:p><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr></w:p>`;
  x=x.replace(/<w:body>/,"<w:body>"+bloc);Z.file("word/document.xml",x);return Z.generateAsync({type:"uint8array"});}
/* ---------- Utilitaires ---------- */
const nomFichier=s=>txt(s).normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^\w.\- ]+/g,"_").replace(/\s+/g,"_").replace(/_+/g,"_").replace(/^_|_$/g,"");
async function telecharger(nom,donnees,type){const b=donnees instanceof Blob?donnees:new Blob([donnees],{type:type||"application/octet-stream"});
  const u=URL.createObjectURL(b),a=document.createElement("a");a.href=u;a.download=nom;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),4000);}
async function zip(fichiers){await zipLib();const Z=new JSZip();for(const [n,d] of fichiers)Z.file(n,d);return Z.generateAsync({type:"blob"});}
window.KepDoc={xlsx,pdf,courriers,qualiopi,qualiopiDocx,zip,telecharger,nomFichier,chargerZip:zipLib};
})();
