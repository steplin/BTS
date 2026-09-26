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

/* ---------- PDF (pdf-lib, polices standard Helvetica) ---------- */
const PAGES={A3:[841.89,1190.55],A4:[595.28,841.89]};
function rgb(h){const n=parseInt(h,16);return PDFLib.rgb(((n>>16)&255)/255,((n>>8)&255)/255,(n&255)/255);}
async function pdf(feuilles,props={}){await pdfLib();const {PDFDocument,StandardFonts,rgb:RGB}=PDFLib;
  const doc=await PDFDocument.create();doc.setTitle(props.titre||"");doc.setCreator("Kerplouz LaSalle");doc.setProducer("keprojects");
  const F={r:await doc.embedFont(StandardFonts.Helvetica),b:await doc.embedFont(StandardFonts.HelveticaBold),i:await doc.embedFont(StandardFonts.HelveticaOblique),bi:await doc.embedFont(StandardFonts.HelveticaBoldOblique)};
  const jeu=new Set(F.r.getCharacterSet());
  const REMP={"\u2018":"'","\u2019":"'","\u201c":'"',"\u201d":'"',"\u00a0":" ","\u202f":" ","\u2009":" ","\u2264":"<=","\u2265":">=","\u2192":"->","\u2190":"<-","\u2212":"-","\u2011":"-","\u2010":"-","\u00ad":"","\t":" "};
  const propre=s=>{let o="";for(const ch of txt(s).normalize("NFC")){if(ch==="\n"){o+=ch;continue;}const cp=ch.codePointAt(0);o+=jeu.has(cp)?ch:(REMP[ch]!==undefined?REMP[ch]:(jeu.has(ch.normalize("NFD").codePointAt(0))?ch.normalize("NFD")[0]:"?"));}return o;};
  const fontDe=x=>x.b&&x.i?F.bi:x.b?F.b:x.i?F.i:F.r;
  const couper=(s,font,sz,larg)=>{const out=[];for(const par of propre(s).split("\n")){if(!par.trim()){out.push("");continue;}let cur="";
      for(const mot of par.split(/ +/)){let m=mot;const essai=cur?cur+" "+m:m;if(font.widthOfTextAtSize(essai,sz)<=larg){cur=essai;continue;}
        if(cur)out.push(cur);cur="";while(font.widthOfTextAtSize(m,sz)>larg&&m.length>1){let k=m.length-1;while(k>1&&font.widthOfTextAtSize(m.slice(0,k),sz)>larg)k--;out.push(m.slice(0,k));m=m.slice(k);}cur=m;}
      out.push(cur);}return out;};
  let totalPages=0;const aNumeroter=[];
  for(const Fe of feuilles){const [W0,H0]=PAGES[Fe.papier||"A4"],[PW,PH]=Fe.paysage?[H0,W0]:[W0,H0];
    const m=Fe.marges||{g:.3,d:.3,h:.9,b:.7},mg=Math.max(22,m.g*72),md=Math.max(22,m.d*72),mh=30,mb=34;
    const largUtile=PW-mg-md,brut=Fe.cols.reduce((a,w)=>a+w*PT_COL,0),k=Math.min(1.25,largUtile/brut),cw=Fe.cols.map(w=>w*PT_COL*k),PAD=3*k+1;
    const X=[mg];cw.forEach(w=>X.push(X[X.length-1]+w));
    const L=Fe.lignes.map(l=>({...l,cells:l.cells.map(x=>x?{...x,sz:(x.sz||10)*k}:x)}));
    // lignes de texte de chaque cellule
    const prep=x=>{if(!x)return;const f=fontDe(x),w=cw.slice(x.c0,x.c0+(x.cs||1)).reduce((a,b)=>a+b,0)-2*PAD;x.f=f;x.L=couper(x.v,f,x.sz,Math.max(8,w));x.lh=x.sz*1.2;x.need=x.L.length*x.lh+2*PAD;};
    L.forEach(l=>l.cells.forEach((x,ci)=>{if(x){x.c0=ci;prep(x);}}));
    const H=L.map(l=>Math.max(l.h?l.h*k*0.9:0,...l.cells.filter(x=>x&&!(x.rs>1)).map(x=>x.need),12*k));
    L.forEach((l,r)=>l.cells.forEach(x=>{if(x&&x.rs>1){const t=H.slice(r,r+x.rs).reduce((a,b)=>a+b,0);if(x.need>t)H[r+x.rs-1]+=x.need-t;}}));
    // blocs ins\u00e9cables (lignes li\u00e9es par une fusion verticale)
    const fin=L.map((_,r)=>r);L.forEach((l,r)=>l.cells.forEach(x=>{if(x&&x.rs>1)for(let q=r;q<r+x.rs;q++)fin[q]=Math.max(fin[q],r+x.rs-1);}));
    const blocs=[];for(let r=0;r<L.length;){let e=fin[r];for(let q=r;q<=e;q++)e=Math.max(e,fin[q]);blocs.push([r,e]);r=e+1;}
    const tete=Fe.repeter?L.slice((Fe.repeterDe||1)-1,(Fe.repeterDe||1)-1+Fe.repeter).map((_,i)=>(Fe.repeterDe||1)-1+i):[];
    const hautPDF=(Fe.hautPDF||[]).map(l=>({...l,cells:l.cells.map((x,ci)=>x?{...x,sz:(x.sz||10)*k,c0:ci}:x)}));hautPDF.forEach(l=>l.cells.forEach(prep));
    const HH=hautPDF.map(l=>Math.max(...l.cells.filter(Boolean).map(x=>x.need),10));
    let page,y;
    const dessinerCellule=(x,x0,y0,w,h,lignesTxt)=>{if(x.fond)page.drawRectangle({x:x0,y:y0-h,width:w,height:h,color:rgb(x.fond)});
      if(x.bord!==false)page.drawRectangle({x:x0,y:y0-h,width:w,height:h,borderColor:RGB(0,0,0),borderWidth:0.6});
      const T=lignesTxt||x.L,th=T.length*x.lh,col=x.coul?rgb(x.coul):RGB(0,0,0);let ty=x.va==="t"?y0-PAD-x.sz:y0-(h-th)/2-x.sz*0.95;
      for(const t of T){if(ty<y0-h+1)break;const tw=x.f.widthOfTextAtSize(t,x.sz);const tx=x.al==="c"?x0+(w-tw)/2:x.al==="r"?x0+w-PAD-tw:x0+PAD;
        if(t)page.drawText(t,{x:tx,y:ty,size:x.sz,font:x.f,color:col});ty-=x.lh;}};
    const dessinerLigne=(l,hl,y0)=>{l.cells.forEach((x,ci)=>{if(!x||x.rs>1)return;const w=cw.slice(ci,ci+(x.cs||1)).reduce((a,b)=>a+b,0);dessinerCellule(x,X[ci],y0,w,hl);});};
    // lignes r0..r1 enti\u00e8res, fusions verticales comprises
    const dessinerPlage=(r0,r1)=>{const Y0={};for(let r=r0;r<=r1;r++){Y0[r]=y;dessinerLigne(L[r],H[r],y);y-=H[r];}
      for(let r=r0;r<=r1;r++)L[r].cells.forEach((x,ci)=>{if(x&&x.rs>1){const h=H.slice(r,Math.min(r1,r+x.rs-1)+1).reduce((s,v)=>s+v,0),w=cw.slice(ci,ci+(x.cs||1)).reduce((s,v)=>s+v,0);dessinerCellule(x,X[ci],Y0[r],w,h);}});};
    let premiere=true,prochaine=0;
    const nouvellePage=()=>{page=doc.addPage([PW,PH]);totalPages++;aNumeroter.push({page,Fe,mg,md,mb,k});y=PH-mh;
      hautPDF.forEach((l,i)=>{l.cells.forEach((x,ci)=>{if(!x)return;const w=cw.slice(ci,ci+(x.cs||1)).reduce((a,b)=>a+b,0);dessinerCellule(x,X[ci],y,w,HH[i]);});y-=HH[i];});
      if(hautPDF.length)y-=6;
      if(!premiere&&tete.length&&(Fe.finRepeter==null||prochaine<Fe.finRepeter))dessinerPlage(tete[0],tete[tete.length-1]);premiere=false;};
    const bas=()=>mb+14;
    nouvellePage();
    const hTete=tete.reduce((s,r)=>s+H[r],0),hHaut=HH.reduce((s,v)=>s+v,0)+(hautPDF.length?6:0);
    for(const [a,e] of blocs){prochaine=a;
      const hb=H.slice(a,e+1).reduce((s,v)=>s+v,0);
      // l'en-t\u00eate du tableau reste avec la premi\u00e8re ligne de donn\u00e9es
      const avecSuite=tete.length&&a===tete[0]?H.slice(e+1,e+2).reduce((s,v)=>s+v,0):0;
      if(y-hb-avecSuite<bas()&&y<PH-mh-hHaut-hTete-8)nouvellePage();
      if(y-hb>=bas()){dessinerPlage(a,e);continue;}
      // lignes du bloc, avec coupure si le bloc d\u00e9passe une page
      const fus=[];for(let r=a;r<=e;r++)L[r].cells.forEach((x,ci)=>{if(x&&x.rs>1)fus.push({x,ci,r0:r,r1:r+x.rs-1,reste:x.L.slice()});});
      let debutSeg=y,rSeg=a;
      const fermerSeg=(rLast,yFin)=>{for(const f of fus){if(f.r1<rSeg||f.r0>rLast)continue;const top=f.r0>=rSeg?yTop(f.r0):debutSeg,bottom=f.r1<=rLast?yBas(f.r1):yFin;
          const h=top-bottom,w=cw.slice(f.ci,f.ci+(f.x.cs||1)).reduce((s,v)=>s+v,0),n=Math.max(0,Math.floor((h-2*PAD)/f.x.lh)),part=f.reste.splice(0,f.r1<=rLast?f.reste.length:n);
          dessinerCellule({...f.x,va:f.r1<=rLast&&f.r0>=rSeg?f.x.va:"t"},X[f.ci],top,w,h,part);}};
      const Y={};const yTop=r=>Y[r],yBas=r=>Y[r]-H[r];
      for(let r=a;r<=e;r++){if(y-H[r]<bas()){fermerSeg(r-1,y);prochaine=r;nouvellePage();debutSeg=y;rSeg=r;}
        Y[r]=y;dessinerLigne(L[r],H[r],y);y-=H[r];}
      fermerSeg(e,y);}
    for(const l of (Fe.apres||[])){const ll={...l,cells:l.cells.map((x,ci)=>x?{...x,sz:(x.sz||10)*k,c0:ci}:x)};ll.cells.forEach(prep);
      const h=Math.max(...ll.cells.filter(Boolean).map(x=>x.need),l.h?l.h*k:0,10);if(y-h<bas())nouvellePage();
      ll.cells.forEach((x,ci)=>{if(!x)return;const w=cw.slice(ci,ci+(x.cs||1)).reduce((s,v)=>s+v,0);dessinerCellule(x,X[ci],y,w,h);});y-=h;}}
  // pieds de page : \u00e9tablissement \u00e0 gauche, n / N \u00e0 droite (num\u00e9rotation par document)
  let num=0,parDoc=new Map();aNumeroter.forEach(p=>parDoc.set(p.Fe,(parDoc.get(p.Fe)||0)+1));const vu=new Map();
  for(const p of aNumeroter){const n=(vu.get(p.Fe)||0)+1;vu.set(p.Fe,n);const sz=9;const {width}=p.page.getSize();
    if(p.Fe.piedPDF)p.page.drawText(propre(p.Fe.piedPDF),{x:p.mg,y:p.mb-12,size:sz,font:F.i});
    const t=n+"/"+parDoc.get(p.Fe);p.page.drawText(t,{x:width-p.md-F.i.widthOfTextAtSize(t,sz),y:p.mb-12,size:sz,font:F.i});num++;}
  return doc.save();}

/* ---------- Utilitaires ---------- */
const nomFichier=s=>txt(s).normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^\w.\- ]+/g,"_").replace(/\s+/g,"_").replace(/_+/g,"_").replace(/^_|_$/g,"");
async function telecharger(nom,donnees,type){const b=donnees instanceof Blob?donnees:new Blob([donnees],{type:type||"application/octet-stream"});
  const u=URL.createObjectURL(b),a=document.createElement("a");a.href=u;a.download=nom;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),4000);}
async function zip(fichiers){await zipLib();const Z=new JSZip();for(const [n,d] of fichiers)Z.file(n,d);return Z.generateAsync({type:"blob"});}
window.KepDoc={xlsx,pdf,zip,telecharger,nomFichier,chargerZip:zipLib};
})();
