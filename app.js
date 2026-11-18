const CONFIG = {
  brand:'MandaGroup', tagline:'Partner Keuangan Rapi',
  entity:'PT Manda Grup Nusantara', email:'mandagrupnusantara@gmail.com', phone:'',
  alamat:'', domain:'mandagrup.web.id', npwp:'',
  bank:'Rekening contoh: Bank CIMB Niaga 1122334455 a.n. PT Manda Grup Nusantara',
  deadlineDay:10, ppn:.11, pph:.005, modalAwal:0
};
const TODAY = new Date(2026,10,13);
const BULAN=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const BLN3=['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
const J=1e6;
const pad2=n=>String(n).padStart(2,'0');
const T=(m,d)=>new Date(2026,m,d);
const periodLabel=t=>BULAN[t-1]+' 2026';
const num=n=>String(Math.round(Math.abs(n))).replace(/\B(?=(\d{3})+(?!\d))/g,'.');
const rp=n=>(n<0?'-':'')+'Rp '+num(n);
function short(n){ const a=Math.abs(n); if(a>=1e9) return 'Rp '+(a/1e9).toFixed(2).replace('.',',')+' M'; if(a>=1e6) return 'Rp '+(a/1e6).toFixed(a<1e8?1:0).replace('.',',')+' jt'; return rp(n); }
const dfmt=d=>d.getDate()+' '+BLN3[d.getMonth()]+' '+d.getFullYear();
const daysTo=d=>Math.round((d-TODAY)/864e5);
function workdays(a,b){ let n=0; const d=new Date(a); while(d<b){ d.setDate(d.getDate()+1); const w=d.getDay(); if(w!==0&&w!==6) n++; } return n; }
function stamp(){ const n=new Date(); return '13 Nov, '+pad2(n.getHours())+':'+pad2(n.getMinutes()); }
const roundK=n=>Math.round(n/1000)*1000;

const DIV={pembukuan:'Pembukuan',pajak:'Perpajakan',review:'Review & Mutu',admin:'Admin & Penagihan'};
const USERS=[
 {id:'md',nama:'Manda',role:'superuser',jabatan:'Managing Director'}
];
const CLIENTS=[
];
const SERVICES=[
 {id:'S01',nama:'Review Pencatatan Keuangan Bulanan',desc:'Pemeriksaan kewajaran dan kelengkapan laporan keuangan yang disusun klien, disertai temuan dan rekomendasi perbaikan.',div:'review',siklus:'Bulanan',harga:1500000,sla:3},
 {id:'S02',nama:'Pembukuan Bulanan',desc:'Pencatatan transaksi dan penyusunan laporan keuangan bulanan oleh tim.',div:'pembukuan',siklus:'Bulanan',harga:2500000,sla:5},
 {id:'S03',nama:'Penyusunan Laporan Keuangan Kuartalan',desc:'Laporan keuangan per kuartal siap untuk kebutuhan bank, investor, atau audit.',div:'pembukuan',siklus:'Kuartalan',harga:4500000,sla:7},
 {id:'S04',nama:'Pendampingan Pajak Bulanan (SPT Masa)',desc:'Perhitungan, pengingat, dan pelaporan SPT Masa.',div:'pajak',siklus:'Bulanan',harga:1500000,sla:3},
 {id:'S05',nama:'Pendampingan SPT Tahunan Badan',desc:'Penyusunan dan pelaporan SPT Tahunan Badan beserta rekonsiliasi fiskal.',div:'pajak',siklus:'Tahunan',harga:6000000,sla:10},
 {id:'S06',nama:'Konsultasi Keuangan dan Pajak',desc:'Sesi konsultasi terjadwal dengan tim atau mitra ahli.',div:'review',siklus:'Per sesi',harga:750000,sla:2}
];
let ENG=[
];
const DOCREQ=[{k:'lr',nama:'Laporan Laba Rugi'},{k:'nr',nama:'Neraca'},{k:'bb',nama:'Buku Besar atau Jurnal Umum'},{k:'rk',nama:'Rekening Koran atau Mutasi Bank'},{k:'pu',nama:'Rekap Piutang dan Utang'}];
const FIG=[['pendapatan','Total pendapatan'],['beban','Total beban (termasuk pajak)'],['laba','Laba bersih'],['aset','Total aset'],['liab','Total liabilitas'],['ekuitas','Total ekuitas'],['kas0','Saldo kas awal'],['masuk','Total penerimaan kas'],['keluar','Total pengeluaran kas'],['kas1','Saldo kas akhir']];
const AREAS=['Neraca','Laba rugi','Arus kas','Piutang usaha','Utang usaha','Persediaan','Aset tetap','Pajak','Kelengkapan dokumen','Lainnya'];

/* ---------- data awal ---------- */
const RV={};
const EXPR={}, ASR={}, KASR={};
function seedFig(cid,t){
  const rev=RV[cid][t]*J, exp=roundK(rev*EXPR[cid]), aset=roundK(RV[cid][Math.min(t,9)]*J*ASR[cid]), liab=roundK(aset*.4);
  const kas1=roundK(RV[cid][Math.max(7,t)]*J*KASR[cid]), kas0=t>7?roundK(RV[cid][t-1]*J*KASR[cid]):roundK(rev*KASR[cid]*.95), masuk=roundK(rev*.97);
  return {pendapatan:rev,beban:exp,laba:rev-exp,aset,liab,ekuitas:aset-liab,kas0,masuk,keluar:kas0+masuk-kas1,kas1};
}
const ALLF={lr:'laba-rugi.xlsx',nr:'neraca.xlsx',bb:'buku-besar.xlsx',rk:'rekening-koran.pdf',pu:'rekap-piutang-utang.xlsx'};
const SUBS=[];
function addSub(o){ const s=Object.assign({id:'P'+(SUBS.length+1),rev:1,files:{},findings:[],hist:[],status:'belum',catatan:'',fig:null,pend:null},o); SUBS.push(s); return s; }
const UNAME=id=>(USERS.find(u=>u.id===id)||{nama:'-'}).nama;

const dueOf=s=>(s.status==='perbaikan'&&s.fixDue)?s.fixDue:T(s.t,CONFIG.deadlineDay);
const engOf=(cid,sid)=>ENG.find(e=>e.cid===cid&&e.sid===sid);
const reviewerOf=s=>{ const e=engOf(s.cid,'S01'); return s.reviewer||(e&&e.pic)||'md'; };
const mitraOf=s=>s.mitra!==undefined?s.mitra:((engOf(s.cid,'S01')||{}).mitra||null);
const slaLeft=s=>{ const sv=SERVICES[0].sla; return sv-workdays(s.sentAt,TODAY); };

/* pra-periksa otomatis */
function prevFig(cid,t){ const p=SUBS.filter(s=>s.cid===cid&&s.t<t&&s.fig).sort((a,b)=>b.t-a.t)[0]; return p?p.fig:null; }
function runChecks(f,files,prev,catatan){
  const n=k=>+f[k]||0, tol=1000, o=[];
  const miss=DOCREQ.filter(d=>!files[d.k]);
  o.push({id:'doc',lvl:miss.length?'err':'ok',t:'Kelengkapan dokumen wajib',d:miss.length?'Belum diunggah: '+miss.map(x=>x.nama).join(', '):'Seluruh dokumen wajib sudah diunggah'});
  if(!n('pendapatan')&&!n('aset')){ o.push({id:'kosong',lvl:'err',t:'Angka utama laporan',d:'Angka laporan belum diisi'}); return o; }
  const d1=n('aset')-n('liab')-n('ekuitas'); o.push({id:'neraca',lvl:Math.abs(d1)<=tol?'ok':'err',t:'Persamaan akuntansi: aset = liabilitas + ekuitas',d:Math.abs(d1)<=tol?'Neraca seimbang':'Tidak seimbang, selisih '+rp(Math.abs(d1))});
  const d2=n('pendapatan')-n('beban')-n('laba'); o.push({id:'laba',lvl:Math.abs(d2)<=tol?'ok':'err',t:'Laba bersih = pendapatan - beban',d:Math.abs(d2)<=tol?'Perhitungan laba sesuai':'Selisih '+rp(Math.abs(d2))});
  const d3=n('kas0')+n('masuk')-n('keluar')-n('kas1'); o.push({id:'kas',lvl:Math.abs(d3)<=tol?'ok':'err',t:'Saldo kas akhir = kas awal + penerimaan - pengeluaran',d:Math.abs(d3)<=tol?'Arus kas sesuai saldo':'Selisih '+rp(Math.abs(d3))});
  if(prev){
    const dk=n('kas0')-prev.kas1; o.push({id:'kontinu',lvl:Math.abs(dk)<=tol?'ok':'warn',t:'Saldo kas awal sama dengan saldo kas akhir periode lalu',d:Math.abs(dk)<=tol?'Berkesinambungan':'Berbeda '+rp(Math.abs(dk))+' dari periode lalu ('+rp(prev.kas1)+')'});
    const ch=prev.pendapatan?(n('pendapatan')-prev.pendapatan)/prev.pendapatan:0; const big=Math.abs(ch)>.4;
    o.push({id:'tren',lvl:big?'warn':'ok',t:'Kewajaran perubahan pendapatan',d:big?'Berubah '+(ch*100).toFixed(0)+'% dibanding periode lalu; reviewer akan menanyakan penyebabnya':'Perubahan '+(ch*100).toFixed(1).replace('.',',')+'% dari periode lalu, dalam batas wajar'});
  }
  return o;
}
const checkSummary=c=>c.some(x=>x.lvl==='err')?'err':c.some(x=>x.lvl==='warn')?'warn':'ok';

const NOTIFS=[]; let NSEQ=0;
function notify(to,text,view,sub){ NSEQ++; NOTIFS.unshift({id:'N'+NSEQ,to,text,view,sub:sub||null,ts:stamp(),read:false}); }
let RULES=[{k:'h7',label:'7 hari sebelum batas',off:-7,on:true,ch:'Portal, email'},{k:'h3',label:'3 hari sebelum batas',off:-3,on:true,ch:'Portal, email, WhatsApp'},{k:'h1',label:'1 hari sebelum batas',off:-1,on:true,ch:'Portal, WhatsApp'},{k:'h0',label:'Pada hari batas',off:0,on:true,ch:'Portal, email, WhatsApp'},{k:'late',label:'Setiap 3 hari setelah terlambat',off:null,on:true,ch:'Portal, email, WhatsApp'}];
const REMLOG=[];
const THREADS={};
const DOCS=[];
const INVOICES=[]; let ISEQ=0;
function mkInv(o){ ISEQ++; const code=o.jenis==='dp'?'-DP':o.jenis==='pelunasan'?'-PL':''; const i=Object.assign({no:'AKT'+code+'-26'+pad2(o.tgl.getMonth()+1)+'-'+String(ISEQ).padStart(3,'0'),ppn:true,status:'terkirim',by:'md',jenis:'penuh',ref:'',dpPct:30,dpRef:''},o); INVOICES.push(i); return i; }
const L=(sid,desc)=>({desc:desc||SERVICES.find(s=>s.id===sid).nama,qty:1,harga:SERVICES.find(s=>s.id===sid).harga});


const MON=m=>BULAN[m-1]+' 2026';
INVOICES.forEach(i=>{ if(i.status==='lunas'&&!i.tglBayar) i.tglBayar=new Date(i.tgl.getTime()+Math.min(10,Math.max(3,Math.round((i.due-i.tgl)/864e5)-3))*864e5); });
function invTot(i){ const kontrak=i.lines.reduce((s,l)=>s+(+l.qty||0)*(+l.harga||0),0), j=i.jenis||'penuh', pct=+i.dpPct||30; let dpp=kontrak, dpBase=0;
  if(j==='dp') dpp=Math.round(kontrak*pct/100);
  else if(j==='pelunasan'){ const ref=i.dpRef?INVOICES.concat(CINV).find(x=>x.no===i.dpRef):null; dpBase=ref?invTot(ref).dpp:Math.round(kontrak*pct/100); dpp=Math.max(0,kontrak-dpBase); }
  const ppn=i.ppn?Math.round(dpp*CONFIG.ppn):0; return {kontrak,dpp,ppn,total:dpp+ppn,dpBase,jenis:j,pct}; }
const invD=d=>({lines:d.lines,ppn:d.ppn,jenis:d.jenis,dpPct:d.dpPct,dpRef:d.dpRef});
const jenisTag=i=>i.jenis==='dp'?' '+pill('info','DP'):i.jenis==='pelunasan'?' '+pill('good','Pelunasan'):'';


const SAMPLE_BRAND={};
CLIENTS.forEach(c=>{ c.brand=Object.assign({logo:null,seq:0,catatan:'Terima kasih atas kepercayaan Anda.'},SAMPLE_BRAND[c.id]); c.finAccess=[...new Set(ENG.filter(e=>e.cid===c.id&&['pembukuan','pajak'].includes(SERVICES.find(x=>x.id===e.sid).div)).map(e=>e.pic))]; });
const CINV=[];
function cinvNo(c,tgl,j){ c.brand.seq++; return 'INV-'+(j==='dp'?'DP-':j==='pelunasan'?'PL-':'')+c.brand.pre+'-'+String(tgl.getFullYear()).slice(2)+pad2(tgl.getMonth()+1)+'-'+String(c.brand.seq).padStart(3,'0'); }
function mkCinv(o){ const c=CLIENTS.find(x=>x.id===o.cid); const i=Object.assign({ppn:false,status:'terkirim',addr:'',email:'',note:'',jenis:'penuh',ref:'',dpPct:30,dpRef:''},o); i.no=i.no||cinvNo(c,i.tgl,i.jenis); CINV.push(i); return i; }

const SVSHORT={S01:'Review',S02:'Pembukuan',S03:'Lap. Kuartalan',S04:'Pajak Bulanan',S05:'SPT Tahunan',S06:'Konsultasi'};
const PAKET=[
 {id:'review',nama:'Review',tag:'Untuk usaha yang sudah punya pembukuan sendiri dan butuh pemeriksa independen.',harga:1500000,svc:['S01'],fitur:['Review laporan keuangan bulanan, hasil maksimal 3 hari kerja','Pra-periksa otomatis dan pengingat batas kirim','Temuan dan rekomendasi perbaikan tertulis','Chat dengan reviewer']},
 {id:'esensial',nama:'Esensial',tag:'Pembukuan dikerjakan tim, lalu direview sebelum Anda terima.',harga:3500000,svc:['S01','S02'],fitur:['Semua fitur paket Review','Pembukuan bulanan oleh tim akuntan','Laporan keuangan bulanan siap pakai','Rekap piutang dan utang']},
 {id:'bisnis',nama:'Bisnis',tag:'Pembukuan dan kepatuhan pajak bulanan dalam satu paket.',harga:4800000,svc:['S01','S02','S04'],fitur:['Semua fitur paket Esensial','Pendampingan SPT Masa bulanan','Pengingat jatuh tempo pajak','Respons chat prioritas satu hari kerja']},
 {id:'korporat',nama:'Korporat',tag:'Untuk usaha yang butuh laporan kuartalan dan konsultasi rutin.',harga:7250000,svc:['S01','S02','S04','S03','S06'],fitur:['Semua fitur paket Bisnis','Laporan keuangan kuartalan','Konsultasi keuangan dan pajak satu sesi per bulan','Dukungan mitra akuntan publik atau konsultan pajak']}
];
const pk=id=>PAKET.find(x=>x.id===id);
const PKSTART={};
CLIENTS.forEach(c=>{ c.paket=PKSTART[c.id]; c.paketNext=null; });

const PAKETLOG=[];
let ESEQ=100;
const pkSvcs=c=>pk(c.paket).svc;
const addons=c=>ENG.filter(e=>e.cid===c.id&&!pkSvcs(c).includes(e.sid));
const monthlyEq=id=>{ const sv=SERVICES.find(x=>x.id===id); return sv.siklus==='Kuartalan'?sv.harga/3:sv.siklus==='Tahunan'?sv.harga/12:sv.harga; };
const paketSaving=p=>Math.round(p.svc.reduce((a,id)=>a+monthlyEq(id),0)-p.harga);
function defPic(cid,sid){ const dv=SERVICES.find(x=>x.id===sid).div; const ex=ENG.find(e=>e.cid===cid&&SERVICES.find(x=>x.id===e.sid).div===dv); if(ex) return ex.pic; const t=USERS.find(x=>x.role==='tim'&&x.div===dv); return t?t.id:'md'; }
function applyPaket(c,id){ const old=pkSvcs(c), nu=pk(id).svc; ENG=ENG.filter(e=>!(e.cid===c.id&&old.includes(e.sid)&&!nu.includes(e.sid))); nu.forEach(sid=>{ if(!ENG.some(e=>e.cid===c.id&&e.sid===sid)) ENG.push({id:'E'+(++ESEQ),cid:c.id,sid,pic:defPic(c.id,sid),mitra:null}); }); c.paket=id; c.paketNext=null; }
function paketLines(c){ const p=pk(c.paket); return [{desc:'Paket '+p.nama+' · layanan akuntansi bulanan '+BULAN[TODAY.getMonth()]+' 2026',qty:1,harga:p.harga}].concat(addons(c).map(e=>({desc:SERVICES.find(x=>x.id===e.sid).nama,qty:1,harga:SERVICES.find(x=>x.id===e.sid).harga}))); }
const FINLOG=[];
const figSubs=cid=>SUBS.filter(s=>s.cid===cid&&s.fig&&s.status!=='belum').sort((a,b)=>a.t-b.t);
const arOf=cid=>CINV.filter(i=>i.cid===cid&&i.status==='terkirim').reduce((a,i)=>a+invTot(i).total,0);

const ICONS={
 home:'<path d="M3 11l9-8 9 8v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/>',
 upload:'<path d="M12 16V4M7 9l5-5 5 5M4 20h16"/>',
 check:'<path d="M20 6L9 17l-5-5"/>',
 folder:'<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
 chat:'<path d="M21 12a8 8 0 0 1-11.6 7.1L3 21l1.9-5.4A8 8 0 1 1 21 12z"/>',
 list:'<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>',
 receipt:'<path d="M6 2h12v20l-3-2-3 2-3-2-3 2zM9 7h6M9 11h6M9 15h3"/>',
 bell:'<path d="M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9M13.7 21a2 2 0 0 1-3.4 0"/>',
 users:'<path d="M17 21v-2a4 4 0 0 0-4-4H7a4 4 0 0 0-4 4v2M10 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8M21 21v-2a4 4 0 0 0-3-3.9M16 3.1a4 4 0 0 1 0 7.8"/>',
 team:'<circle cx="12" cy="8" r="3"/><path d="M5 20v-1a5 5 0 0 1 5-5h4a5 5 0 0 1 5 5v1M4 10a2 2 0 1 0 0-4M20 10a2 2 0 1 1 0-4"/>',
 clock:'<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
 gear:'<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>',
 plus:'<path d="M12 5v14M5 12h14"/>',
 x:'<path d="M18 6L6 18M6 6l12 12"/>',
 copy:'<rect x="9" y="9" width="12" height="12" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h8"/>',
 back:'<path d="M15 18l-6-6 6-6"/>',
 alert:'<path d="M10.3 3.9L1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
 refresh:'<path d="M21 12a9 9 0 0 0-15-6.7L3 8M3 3v5h5M3 12a9 9 0 0 0 15 6.7L21 16M16 16h5v5"/>',
 activity:'<path d="M22 12h-4l-3 9L9 3l-3 9H2"/>',
 inbox:'<path d="M22 12h-6l-2 3h-4l-2-3H2"/><path d="M5.5 5.1L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.5-6.9A2 2 0 0 0 16.7 4H7.2a2 2 0 0 0-1.7 1.1z"/>',
 hourglass:'<path d="M5 22h14M5 2h14M17 22v-4.2a2 2 0 0 0-.6-1.4L12 12l-4.4 4.4a2 2 0 0 0-.6 1.4V22M7 2v4.2a2 2 0 0 0 .6 1.4L12 12l4.4-4.4A2 2 0 0 0 17 6.2V2"/>',
 filecheck:'<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7z"/><path d="M14 2v5h6M9 15l2 2 4-4"/>',
 wallet:'<path d="M19 7V5a2 2 0 0 0-2-2H5a2 2 0 0 0 0 4h14a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5"/><path d="M16 13h.01"/>',
 send:'<path d="M22 2L11 13M22 2l-7 20-4-9-9-4z"/>',
 arrow:'<path d="M5 12h14M13 6l6 6-6 6"/>',
 trend:'<path d="M22 7l-9 9-4-4-7 7M16 7h6v6"/>',
 sun:'<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>',
 moon:'<path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/>',
 shield:'<path d="M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6z"/><path d="M9 12l2 2 4-4"/>',
 lock:'<rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>',
 out:'<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9"/>',
 eye:'<path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7S1 12 1 12z"/><circle cx="12" cy="12" r="3"/>'
};
const ic=n=>'<svg class="ic" viewBox="0 0 24 24" aria-hidden="true">'+ICONS[n]+'</svg>';
const esc=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const $=id=>document.getElementById(id);

const S={kt:'ringkas',ktp:10,kxd:null,theme:'auto',cdt:'ringkas',pkf:null,cinvF:'semua',cin:null,fkc:null,fkp:null,fkpc:null,ftab:'lapor',fst:'lr',entered:false,uid:'md',view:'dashboard',sub:null,cd:null,modal:null,msgCid:null,docCid:null,qF:'semua',invF:'semua',ltab:'katalog',form:null,fnd:{area:'Neraca',sev:'sedang',teks:''},note:'',draft:'',reply:{},invd:null,nc:{nama:'',bentuk:'PT',usaha:'',email:''}};
const me=()=>USERS.find(u=>u.id===S.uid);
const cl=id=>CLIENTS.find(c=>c.id===id);
const svc=id=>SERVICES.find(s=>s.id===id);
const isSU=()=>me().role==='superuser', isTim=()=>me().role==='tim', isMitra=()=>me().role==='mitra', isKlien=()=>me().role==='klien';
const isAdminDiv=()=>isSU()||(isTim()&&me().div==='admin');

function mySubs(){
  const u=me();
  if(u.role==='superuser') return SUBS.slice();
  if(u.role==='klien') return SUBS.filter(s=>s.cid===u.cid);
  if(u.role==='mitra') return SUBS.filter(s=>mitraOf(s)===u.id);
  return SUBS.filter(s=>myClients().some(c=>c.id===s.cid));
}
function myClients(){
  const u=me();
  if(u.role==='superuser') return CLIENTS.slice();
  if(u.role==='klien') return CLIENTS.filter(c=>c.id===u.cid);
  if(u.role==='mitra'){ const ids=new Set(SUBS.filter(s=>mitraOf(s)===u.id).map(s=>s.cid)); ENG.filter(e=>e.mitra===u.id).forEach(e=>ids.add(e.cid)); return CLIENTS.filter(c=>ids.has(c.id)); }
  if(u.div==='admin'||u.div==='review') return CLIENTS.slice();
  const ids=new Set(ENG.filter(e=>e.pic===u.id).map(e=>e.cid)); return CLIENTS.filter(c=>ids.has(c.id));
}
function canReview(s){ const u=me(); return u.role==='superuser'||(u.role==='tim'&&(u.div==='review'||reviewerOf(s)===u.id))||(u.role==='mitra'&&mitraOf(s)===u.id); }
const canRelease=()=>isSU()||(isTim()&&me().div==='review');
const myNotifs=()=>{ const u=me(); const k=u.role==='klien'?'c:'+u.cid:'u:'+u.id; return NOTIFS.filter(n=>n.to===k); };
const unreadN=()=>myNotifs().filter(n=>!n.read).length;
function unreadMsgs(cid){ const mine=isKlien()?'klien':'tim'; return (THREADS[cid]||[]).filter(m=>m.side!==mine&&!m.read).length; }


const pill=(c,t)=>'<span class="pill '+c+'">'+t+'</span>';
function subPill(s,client){
  const st=s.status;
  if(st==='belum'){ const d=daysTo(dueOf(s)); return d<0?pill('bad','Terlambat '+(-d)+' hari'):s.draft?pill('mute','Draft tersimpan'):pill('mute','Belum dikirim'); }
  if(st==='menunggu') return pill('info','Menunggu review');
  if(st==='review') return pill('info','Sedang direview');
  if(st==='qc') return client?pill('info','Sedang direview'):pill('warn','Menunggu rilis mutu');
  if(st==='perbaikan') return pill('bad','Perlu perbaikan');
  return pill('good','Sesuai');
}
const sevPill=v=>v==='tinggi'?pill('bad','Tinggi'):v==='sedang'?pill('warn','Sedang'):pill('mute','Rendah');
const fndPill=f=>f.st==='ditutup'?pill('good','Ditutup'):f.st==='dijawab'?pill('info','Sudah ditanggapi'):pill('bad','Perlu ditanggapi');
function invStatus(i){ if(i.status==='lunas') return pill('good','Lunas'); if(i.status==='draft') return pill('mute','Draft'); if(i.status==='verifikasi') return pill('info','Menunggu verifikasi'); const d=daysTo(i.due); return d<0?pill('bad','Terlambat '+(-d)+' hari'):pill('warn','Jatuh tempo '+dfmt(i.due)); }
const tile=(k,v,d)=>'<div class="card tile"><span class="k">'+k+'</span><span class="v num">'+v+'</span><span class="d">'+d+'</span></div>';
const head=(t,s,r)=>'<div class="page-h"><div><h1>'+t+'</h1>'+(s?'<p>'+s+'</p>':'')+'</div><div class="grow"></div>'+(r||'')+'</div>';


const THEME_KEY='akt-theme';
function themePref(){ try{ return localStorage.getItem(THEME_KEY)||'auto'; }catch(e){ return 'auto'; } }
function isDark(){ if(S.theme==='dark') return true; if(S.theme==='light') return false; try{ return !!(window.matchMedia&&window.matchMedia('(prefers-color-scheme: dark)').matches); }catch(e){ return false; } }
function applyTheme(){ const r=document.documentElement; if(S.theme==='auto') r.removeAttribute('data-theme'); else r.setAttribute('data-theme',S.theme); }
const themeBtn=cls=>'<button class="'+cls+'" data-a="theme" aria-label="Ganti ke mode '+(isDark()?'terang':'gelap')+'" title="Mode '+(isDark()?'terang':'gelap')+'">'+ic(isDark()?'sun':'moon')+'</button>';
const logoMark=(n,tile)=>{ const t=tile||'#F5EDE6',l='#7A2418',r='#B04A2C'; return ''; };
const wordmark=()=>'Manda<i>Group</i>';
const DEMO_PW='demo1234';
const demoEmail=u=>u.id==='md'?CONFIG.email:String(u.nama).split(/[\s,]+/)[0].toLowerCase()+'@mandagroup.example';
function landing(){
  const lg=S.lg, pts=[['trend','Laporan otomatis dan transparan','Neraca, laba rugi, dan arus kas tersusun dari transaksi Anda, lengkap dengan proyeksi.'],['lock','Dokumen tertata dan aman','Berkas, invoice, dan catatan keuangan tersimpan rapi dengan akses terbatas per klien.'],['team','Didampingi tim profesional','Reviewer dan konsultan kami memeriksa laporan serta memberi rekomendasi yang jelas.']];
  const grp=(t,list)=>'<div class="demo-grp"><div class="upper">'+t+'</div>'+list.map(u=>'<button type="button" class="demo-acct" data-a="lgfill" data-id="'+u.id+'"><span><b>'+esc(u.role==='klien'?cl(u.cid).nama:u.nama)+'</b><span class="muted">'+esc(demoEmail(u))+'</span></span><span class="muted">'+esc(u.role==='klien'?u.nama:(u.jabatan||roleLabel(u)))+'</span></button>').join('')+'</div>';
  return '<div class="login">'+themeBtn('bell theme-fab')+
  '<div class="login-art"><div class="login-brand"><div class="lb-row">'+logoMark(44)+'<b>'+wordmark()+'</b></div><div class="brand-sub">'+esc(CONFIG.tagline)+'</div></div>'+
  '<div class="login-welcome"><h2>Keuangan usaha Anda, tertata dan terpantau.</h2><p>Kirim data transaksi, pantau laporan yang tersusun otomatis, dan terima rekomendasi konsultan dalam satu portal.</p><ul class="login-points">'+pts.map(x=>'<li>'+ic(x[0])+'<span><b>'+x[1]+'</b>'+x[2]+'</span></li>').join('')+'</ul></div>'+
  '<div class="login-foot"><span>© 2026 '+esc(CONFIG.entity)+'</span>'+(CONFIG.phone?'<span>'+esc(CONFIG.phone)+'</span>':'')+'<span>'+esc(CONFIG.email)+'</span><span>'+esc(CONFIG.domain)+'</span></div></div>'+
  '<div class="login-form"><div class="login-box"><div class="login-logo-mobile">'+logoMark(36)+'<b>'+wordmark()+'</b></div><span class="eyebrow">SELAMAT DATANG</span><h1 class="login-title">Masuk ke '+CONFIG.brand+'</h1><p class="note" style="margin:8px 0 0">Silakan masuk menggunakan email dan kata sandi yang Anda terima dari kami.</p>'+
  '<label class="f" style="margin-top:22px">Email<input type="email" autocomplete="username" data-input="lg" data-k="email" value="'+esc(lg.email)+'" placeholder="Alamat email Anda"></label>'+
  '<label class="f" style="margin-top:14px">Kata sandi<span class="pw"><input type="'+(lg.show?'text':'password')+'" autocomplete="current-password" data-input="lg" data-k="pw" value="'+esc(lg.pw)+'" placeholder="••••••••"><button type="button" class="btn sm" data-a="lgshow" aria-label="Tampilkan kata sandi">'+ic('eye')+'</button></span></label>'+
  '<div class="row" style="justify-content:space-between;margin-top:10px"><span class="note" style="display:inline-flex;gap:6px;align-items:center">'+ic('shield')+'Koneksi aman</span><button type="button" class="link" data-a="lgforgot">Lupa kata sandi?</button></div>'+
  (lg.err?'<div class="banner bad" role="alert" style="margin-top:14px"><div>'+esc(lg.err)+'</div></div>':'')+
  '<button type="button" class="btn primary lg-btn" data-a="lgsubmit">Masuk</button>'+
  '<div class="demo-box"><div class="upper" style="color:var(--good)">'+ic('shield')+' Akun demo (klik untuk mengisi)</div>'+grp('Klien',USERS.filter(x=>x.role==='klien'))+grp('Pemilik dan tim',USERS.filter(x=>x.role==='superuser'||x.role==='tim'))+grp('Mitra pekerjaan',USERS.filter(x=>x.role==='mitra'))+'<p class="note" style="margin:10px 0 0">Kata sandi demo: '+DEMO_PW+'. Versi produksi memakai email dan kata sandi masing-masing.</p></div></div></div></div>';
}
function navFor(u){
  if(u.role==='klien') return [['dashboard','home','Beranda'],['kirim','upload','Kirim Laporan'],['hasil','check','Hasil Review'],['akun','trend','Laporan & Analisis'],['cinv','receipt','Invoice Usaha'],['usaha','gear','Profil Usaha'],['dokumen','folder','Dokumen'],['pesan','chat','Pesan'],['layanan','list','Layanan Saya'],['tagihan','wallet','Tagihan MandaGroup'],['notif','bell','Notifikasi']];
  if(u.role==='mitra') return [['dashboard','home','Pekerjaan Saya'],['antrean','check','Daftar Review']].concat(finClients().length?[['keuangan','wallet','Keuangan Klien'],['akun','trend','Akuntansi Klien']]:[],[['notif','bell','Notifikasi']]);
  const n=[['dashboard','home','Dasbor'],['klien','users',u.role==='superuser'?'Klien':'Klien Saya'],['antrean','check','Antrean Review']];
  if(u.role==='superuser'){ n.push(['layanan','list','Layanan dan Penugasan']); n.push(['tim','team','Tim dan Mitra']); }
  if(isAdminDiv()) n.push(['invoice','receipt','Invoice']);
  if(u.role==='superuser') n.push(['kakt','wallet','Keuangan MandaGroup']);
  if(u.role==='superuser'||finClients().length) n.push(['keuangan','wallet','Keuangan Klien'],['akun','trend','Akuntansi Klien']);
  n.push(['pengingat','clock','Pengingat'],['dokumen','folder','Dokumen'],['pesan','chat','Pesan'],['notif','bell','Notifikasi']);
  if(u.role==='superuser') n.push(['pengaturan','gear','Pengaturan']);
  return n;
}
function badgeFor(k){
  const u=me();
  if(k==='notif') return unreadN();
  if(k==='pesan') return myClients().reduce((s,c)=>s+unreadMsgs(c.id),0);
  if(k==='antrean'&&!isKlien()) return mySubs().filter(s=>(s.status==='menunggu'||(s.status==='qc'&&canRelease()))&&(canReview(s)||canRelease())).length;
  if(k==='kirim'&&isKlien()) return mySubs().filter(s=>s.status==='belum'||s.status==='perbaikan').length;
  if(k==='tagihan'&&isKlien()) return INVOICES.filter(i=>i.cid===u.cid&&i.status==='terkirim').length;
  return 0;
}
function roleLabel(u){ return u.role==='superuser'?'Superuser':u.role==='tim'?'Tim · '+DIV[u.div]:u.role==='mitra'?'Mitra pekerjaan':'Klien'; }
const NAVG={dashboard:'RUANG KERJA',antrean:'RUANG KERJA',pengingat:'RUANG KERJA',pesan:'RUANG KERJA',kirim:'LAPORAN',hasil:'LAPORAN',klien:'KLIEN & LAYANAN',layanan:'KLIEN & LAYANAN',tim:'KLIEN & LAYANAN',dokumen:'KLIEN & LAYANAN',tagihan:'KEUANGAN',invoice:'KEUANGAN',kakt:'KEUANGAN',keuangan:'KEUANGAN',akun:'KEUANGAN',pengaturan:'SISTEM'};
function shell(){
  const u=me(); const items=navFor(u);
  if(!items.some(i=>i[0]===S.view)) S.view='dashboard';
  const cur=items.find(i=>i[0]===S.view);
  const groups=[]; items.filter(i=>i[0]!=='notif').forEach(i=>{ const KG={dashboard:'PORTAL KLIEN',kirim:'LAPORAN',hasil:'LAPORAN',akun:'LAPORAN',cinv:'KEUANGAN USAHA',usaha:'KEUANGAN USAHA',dokumen:'LAYANAN MandaGroup',pesan:'LAYANAN MandaGroup',layanan:'LAYANAN MandaGroup',tagihan:'LAYANAN MandaGroup'}; const g=u.role==='klien'?(KG[i[0]]||'LAINNYA'):u.role==='mitra'?'MITRA PEKERJAAN':(NAVG[i[0]]||'LAINNYA'); let x=groups.find(y=>y.l===g); if(!x){ x={l:g,i:[]}; groups.push(x); } x.i.push(i); });
  const nav=groups.map(g=>'<div class="nav-group"><div class="nav-label">'+g.l+'</div>'+g.i.map(([k,i,l])=>{ const b=badgeFor(k); return '<button data-a="go" data-v="'+k+'" '+(S.view===k?'aria-current="page"':'')+'>'+ic(i)+l+(b?'<span class="badge">'+b+'</span>':'')+'</button>'; }).join('')+'</div>').join('');
  const ctx=u.role==='klien'?cl(u.cid).nama:u.jabatan;
  const title=cur?cur[2]:'Dasbor';
  return '<aside class="side"><div class="brand">'+logoMark(38)+'<div><b>'+wordmark()+'</b><span>'+CONFIG.tagline+'</span></div></div><nav class="nav" aria-label="Menu utama">'+nav+'</nav>'+
   '<div class="side-foot"><div class="me"><span class="av">'+ini(u.nama)+'</span><span><b>'+esc(u.nama)+'</b><small>'+esc(roleLabel(u))+'</small></span></div><button class="btn sm out" data-a="logout">'+ic('out')+'Keluar</button></div></aside>'+
   '<div class="main"><header class="top"><div class="who"><h1 class="topbar-title">'+esc(title)+'</h1><small class="crumbs">'+CONFIG.brand+' › '+esc(title)+' · '+esc(ctx)+'</small></div><div class="grow"></div>'+themeBtn('bell')+'<button class="bell" data-a="go" data-v="notif" aria-label="Notifikasi">'+ic('bell')+(unreadN()?'<span class="dot">'+unreadN()+'</span>':'')+'</button></header>'+
   '<main class="content dashboard" id="content"><div class="page">'+viewHtml()+'</div></main></div>'+modalHtml();
}
function viewHtml(){
  if(S.sub){ const s=SUBS.find(x=>x.id===S.sub); if(s) return vDetail(s); S.sub=null; }
  const v=S.view, k=isKlien();
  switch(v){
    case 'dashboard': return k?dashKlien():isMitra()?dashMitra():dashStaff();
    case 'kirim': return vKirim();
    case 'hasil': return vHasil();
    case 'antrean': return vAntrean();
    case 'klien': return S.cd?vKlienDetail(cl(S.cd)):vKlien();
    case 'layanan': return k?vLayananKlien():vLayanan();
    case 'tim': return vTim();
    case 'invoice': return vInvoice();
    case 'tagihan': return vTagihan();
    case 'cinv': return vCinv();
    case 'usaha': return vUsaha();
    case 'keuangan': return vKeuangan();
    case 'kakt': return vKakt();
    case 'akun': return vAkun();
    case 'pengingat': return vPengingat();
    case 'dokumen': return vDokumen();
    case 'pesan': return vPesan();
    case 'notif': return vNotif();
    case 'pengaturan': return vPengaturan();
  }
  return '';
}


function actionsKlien(c){
  const a=[];
  SUBS.filter(s=>s.cid===c.id).forEach(s=>{
    const d=daysTo(dueOf(s));
    if(s.status==='belum') a.push({sev:d<0?'bad':d<=7?'warn':'info',t:'Kirim laporan keuangan '+periodLabel(s.t),s:d<0?'Terlambat '+(-d)+' hari (batas '+dfmt(dueOf(s))+')':'Batas kirim '+dfmt(dueOf(s)),v:'kirim'});
    if(s.status==='perbaikan'){ const o=s.findings.filter(f=>f.st==='terbuka').length; a.push({sev:'bad',t:o?o+' temuan '+periodLabel(s.t)+' perlu ditanggapi':'Kirim ulang laporan '+periodLabel(s.t),s:'Batas perbaikan '+dfmt(dueOf(s)),v:'hasil',sub:s.id}); }
  });
  INVOICES.filter(i=>i.cid===c.id&&i.status==='terkirim').forEach(i=>{ const d=daysTo(i.due); a.push({sev:d<0?'bad':'info',t:'Bayar invoice '+i.no+' ('+rp(invTot(i).total)+')',s:d<0?'Terlambat '+(-d)+' hari':'Jatuh tempo '+dfmt(i.due),v:'tagihan'}); });
  const u=unreadMsgs(c.id); if(u) a.push({sev:'info',t:u+' pesan baru dari tim',s:'Buka percakapan',v:'pesan'});
  return a.sort((x,y)=>({bad:0,warn:1,info:2}[x.sev]-{bad:0,warn:1,info:2}[y.sev]));
}
function itemBtn(a){ return '<button class="item itemlink stripe '+a.sev+'" data-a="go" data-v="'+a.v+'"'+(a.sub?' data-sub="'+a.sub+'"':'')+'><div class="grow"><strong>'+esc(a.t)+'</strong><small>'+esc(a.s)+'</small></div></button>'; }
const stepsFor=s=>{ const st=s.status; const a=st==='belum'?['','','']:st==='menunggu'?['done','now','']:(st==='review'||st==='qc')?['done','done now','']:st==='perbaikan'?['done','done','bad']:['done','done','done'];
  const lab=['Dikirim','Direview',st==='perbaikan'?'Perlu perbaikan':st==='disetujui'?'Sesuai':'Hasil'];
  return '<div class="steps">'+lab.map((l,i)=>'<div class="s '+a[i].replace('done now','now')+(a[i]==='done now'?' done':'')+'">'+l+'</div>').join('')+'</div>'; };


const ini=n=>String(n||'?').split(/\s+/).slice(0,2).map(w=>w[0]).join('').toUpperCase();
const PROG={belum:8,menunggu:30,review:55,qc:78,perbaikan:50,disetujui:100};
const STLAB={belum:'Belum dikirim',menunggu:'Menunggu review',review:'Sedang direview',qc:'Kontrol mutu',perbaikan:'Perlu perbaikan',disetujui:'Sesuai'};
const STCOL={belum:'#BFA79B',menunggu:'#D9884F',review:'var(--accent)',qc:'var(--brand)',perbaikan:'#B3261E',disetujui:'#1B7A4B'};
function tsKey(ts){ const m=/^(\d+)\s(\S+),\s(\d+):(\d+)/.exec(ts||''); if(!m) return 0; return (BLN3.indexOf(m[2])*31+(+m[1]))*1440+(+m[3])*60+(+m[4]); }
function actFeed(list,client){ const a=[]; list.forEach(s=>s.hist.forEach(h=>{ if(client&&h.internal) return; a.push({k:tsKey(h.ts),ts:h.ts,who:h.who,text:h.text,c:cl(s.cid).singkat||cl(s.cid).nama}); })); return a.sort((x,y)=>y.k-x.k).slice(0,6); }
const bar=(v,col)=>'<div class="bar"><span style="width:'+v+'%;'+(col?'background:'+col:'')+'"></span></div>';
function intro(sub){ const u=me(), h=new Date().getHours(); const g=h<11?'Selamat pagi':h<15?'Selamat siang':h<18?'Selamat sore':'Selamat malam';
  return '<div class="dashboard-intro"><div><span class="eyebrow">'+CONFIG.brand+' · '+roleLabel(u)+'</span><h1>'+g+', '+esc(u.nama.split(' ')[0])+'<span class="greeting-dot">.</span></h1><p>'+sub+'</p></div><div class="dashboard-date"><span>'+dfmt(TODAY)+'</span><button class="btn sm" data-a="refresh">'+ic('refresh')+'Perbarui</button></div></div>'; }
function focusCard(o){ return '<section class="today-focus"><div class="focus-head"><span class="eyebrow">'+o.eb+'</span>'+ic('activity')+'</div><div class="focus-body"><div class="focus-copy"><h2>'+o.h+'</h2><div class="next-person">'+o.t+'</div><p>'+o.p+'</p>'+(o.cta||'')+'</div>'+(o.side||'')+'</div><div class="focus-footer">'+o.f.map(x=>'<span>'+x+'</span>').join('')+'</div></section>'; }
const quick=items=>'<section class="quick-actions"><span class="eyebrow">AKSES CEPAT</span>'+items.map(i=>'<button class="qa" '+i[0]+'><div class="quick-icon">'+ic(i[1])+'</div><span><b>'+i[2]+'</b><small>'+i[3]+'</small></span><span class="quick-plus">→</span></button>').join('')+'</section>';
const stat=(tone,icon,label,val,note,attrs)=>'<button class="stat-link" '+(attrs||'')+'><div class="stat '+tone+'"><div class="ico">'+ic(icon)+'</div><div class="lbl">'+label+'</div><div class="val num">'+val+'</div><span class="note">'+note+'</span></div></button>';
const feed=(list)=>'<section class="card card-pad activity-card"><span class="eyebrow">AKTIVITAS TERKINI</span><h2>Yang baru di portal</h2><div class="activity-timeline">'+(list.length?list.map(a=>'<div class="activity-item"><span class="activity-icon">'+ic('activity')+'</span><div><b>'+esc(a.text)+'</b><small>'+esc(a.c)+' · '+esc(a.who)+' · '+esc(a.ts)+'</small></div></div>').join(''):'<p class="note">Aktivitas tampil setelah ada data.</p>')+'</div></section>';
const foot=()=>'<div class="workspace-footer"><span>'+esc(CONFIG.entity)+' · '+esc(CONFIG.email)+(CONFIG.phone?' · '+esc(CONFIG.phone):'')+'</span><span>Hasil review dalam '+svc('S01').sla+' hari kerja sejak dokumen lengkap</span></div>';
const ring=(pct,lab)=>{ const r=52,c=2*Math.PI*r; return '<div class="day-progress" role="img" aria-label="Progres '+pct+' persen"><svg viewBox="0 0 124 124"><circle cx="62" cy="62" r="'+r+'" fill="none" stroke="rgba(255,255,255,.18)" stroke-width="9"/><circle class="progress-ring" cx="62" cy="62" r="'+r+'" fill="none" stroke="#D9884F" stroke-width="9" stroke-linecap="round" stroke-dasharray="'+c.toFixed(1)+'" stroke-dashoffset="'+(c*(1-pct/100)).toFixed(1)+'" transform="rotate(-90 62 62)"/></svg><div><b>'+pct+'<small>%</small></b><span>'+lab+'</span></div></div>'; };

function dashKlien(){
  const u=me(), c=cl(u.cid), acts=actionsKlien(c), subs=SUBS.filter(s=>s.cid===c.id).sort((a,b)=>b.t-a.t), cur=subs[0];
  const openF=subs.reduce((n,s)=>n+s.findings.filter(f=>f.st==='terbuka').length,0);
  const inv=INVOICES.filter(i=>i.cid===c.id&&i.status==='terkirim').reduce((s,i)=>s+invTot(i).total,0);
  const per=periodLabel(cur.t), late=cur.status==='belum'&&daysTo(dueOf(cur))<0;
  let t,p,cta;
  if(cur.status==='disetujui'){ t='Laporan '+per+' dinyatakan sesuai'; p='Terima kasih. Periode berikutnya dibuka 1 Desember 2026.'; cta='<button class="focus-time" data-a="go" data-v="hasil" data-sub="'+cur.id+'">'+ic('check')+'Lihat hasil review</button>'; }
  else if(cur.status==='perbaikan'){ t='Laporan '+per+' perlu diperbaiki'; p='Tanggapi temuan reviewer lalu kirim ulang sebelum '+dfmt(dueOf(cur))+'.'; cta='<button class="focus-time" data-a="go" data-v="hasil" data-sub="'+cur.id+'">'+ic('alert')+'Lihat temuan<span>'+openF+' terbuka</span></button>'; }
  else if(cur.status==='belum'){ t='Laporan '+per+' belum dikirim'; p=(late?'Batas kirim '+dfmt(dueOf(cur))+' sudah lewat '+(-daysTo(dueOf(cur)))+' hari.':'Batas kirim '+dfmt(dueOf(cur))+'.'); cta='<button class="focus-time" data-a="go" data-v="kirim">'+ic('upload')+'Kirim laporan</button>'; }
  else { t='Laporan '+per+' sedang kami periksa'; p='Estimasi hasil '+Math.max(0,slaLeft(cur))+' hari kerja lagi.'; cta='<button class="focus-time" data-a="go" data-v="hasil" data-sub="'+cur.id+'">'+ic('clock')+'Lihat status</button>'; }
  return intro(esc(c.nama)+' · review pencatatan keuangan')+
  '<div class="workspace-overview">'+focusCard({eb:'STATUS LAPORAN TERKINI',h:'Periode '+per,t:t,p:p,cta:cta,side:ring(PROG[cur.status],'progres'),f:[subs.filter(s=>s.status==='disetujui').length+' dari '+subs.length+' periode sesuai',STLAB[cur.status]]})+
  quick([['data-a="go" data-v="kirim"','upload','Kirim laporan','Isi angka dan lampirkan dokumen'],['data-a="go" data-v="hasil"','check','Hasil review','Temuan dan status tiap periode'],['data-a="go" data-v="cinv"','receipt','Invoice usaha','Buat dan kirim ke pelanggan Anda'],['data-a="go" data-v="pesan"','chat','Chat dengan tim','Tanya langsung ke reviewer']])+'</div>'+
  '<div class="grid g4 dashboard-stats">'+stat('white','hourglass','Periode terkini',per,subPill(cur,true),'data-a="go" data-v="hasil" data-sub="'+cur.id+'"')+stat('mint','filecheck','Laporan sesuai',subs.filter(s=>s.status==='disetujui').length+' dari '+subs.length,'Periode Agu sampai Okt 2026','data-a="go" data-v="hasil"')+stat(openF?'peach':'white','alert','Temuan terbuka',String(openF),openF?'Perlu tanggapan Anda':'Tidak ada','data-a="go" data-v="hasil"')+stat('blue','receipt','Tagihan belum dibayar',inv?short(inv):'Lunas',inv?'Termasuk PPN':'Tidak ada tagihan','data-a="go" data-v="tagihan"')+'</div>'+
  paketCard(c)+catatanCard(c)+(()=>{ const L=finLatest([c.id]); return '<div class="performance-head" style="margin:-4px 0 -6px"><div><span class="eyebrow">RINGKASAN KEUANGAN USAHA ANDA</span></div><button class="link" data-a="go" data-v="cinv">Invoice usaha</button></div>'+finStrip([['white','trend','Pendapatan',short(L.pend),'Periode terakhir dilaporkan'],['mint','wallet','Laba bersih',short(L.laba),L.pend?(L.laba/L.pend*100).toFixed(1).replace('.',',')+'% dari pendapatan':''],['blue','wallet','Saldo kas akhir',short(L.kas),'Sesuai laporan Anda'],['white','receipt','Piutang pelanggan',short(L.ar),'Invoice usaha belum dibayar','data-a="go" data-v="cinv"']]); })()+
  '<div class="dashboard-detail-grid"><section class="card card-pad"><div class="performance-head"><div><span class="eyebrow">PERLU TINDAKAN</span><h2>Yang perlu Anda selesaikan</h2></div>'+pill('mute',String(acts.length))+'</div>'+(acts.length?'<div class="list" style="margin-top:12px">'+acts.map(itemBtn).join('')+'</div>':'<div class="empty">Tidak ada yang perlu ditindaklanjuti.</div>')+'</section>'+
  '<section class="card card-pad"><div class="performance-head"><div><span class="eyebrow">RIWAYAT</span><h2>Periode laporan</h2></div></div><div class="deadline-list">'+subs.map(x=>'<button class="deadline-row" data-a="go" data-v="hasil" data-sub="'+x.id+'"><div class="dl-date"><b>'+T(x.t,CONFIG.deadlineDay).getDate()+'</b><span>'+BLN3[x.t-1]+'</span></div><div class="grow" style="min-width:0"><b class="trunc">'+periodLabel(x.t)+'</b><small>Batas kirim '+dfmt(T(x.t,CONFIG.deadlineDay))+'</small>'+bar(PROG[x.status],STCOL[x.status])+'</div>'+subPill(x,true)+'</button>').join('')+'</div></section></div>'+foot();
}
function dashStaff(){
  const u=me(), subs=mySubs(), su=isSU();
  const q=subs.filter(s=>s.status==='menunggu'), rv=subs.filter(s=>s.status==='review'), qc=subs.filter(s=>s.status==='qc'), late=subs.filter(s=>(s.status==='belum'&&daysTo(dueOf(s))<0)), perb=subs.filter(s=>s.status==='perbaikan');
  const att=[];
  subs.filter(s=>s.sentAt&&(s.status==='menunggu'||s.status==='review')&&slaLeft(s)<0).forEach(s=>att.push({sev:'bad',t:cl(s.cid).nama+' · '+periodLabel(s.t)+' melewati SLA review',s:'Terlambat '+(-slaLeft(s))+' hari kerja · reviewer '+UNAME(mitraOf(s)||reviewerOf(s)),sub:s.id}));
  qc.forEach(s=>att.push({sev:'warn',t:cl(s.cid).nama+' · hasil mitra menunggu rilis mutu',s:'Dikerjakan '+UNAME(mitraOf(s)),sub:s.id}));
  late.forEach(s=>att.push({sev:'bad',t:cl(s.cid).nama+' belum mengirim laporan '+periodLabel(s.t),s:'Terlambat '+(-daysTo(dueOf(s)))+' hari',v:'pengingat'}));
  q.forEach(s=>att.push({sev:'info',t:cl(s.cid).nama+' · '+periodLabel(s.t)+' menunggu review',s:'Dikirim '+dfmt(s.sentAt),sub:s.id}));
  if(isAdminDiv()) INVOICES.filter(i=>i.status==='verifikasi').forEach(i=>att.push({sev:'info',t:'Verifikasi pembayaran '+i.no+' · '+cl(i.cid).nama,s:'Klien sudah mengonfirmasi transfer',v:'invoice'}));
  if(isAdminDiv()) INVOICES.filter(i=>i.status==='terkirim'&&daysTo(i.due)<0).forEach(i=>att.push({sev:'bad',t:'Invoice '+i.no+' terlambat · '+cl(i.cid).nama,s:'Jatuh tempo '+dfmt(i.due),v:'invoice'}));
  const order={bad:0,warn:1,info:2}; att.sort((a,b)=>order[a.sev]-order[b.sev]);
  const out=INVOICES.filter(i=>i.status==='terkirim'||i.status==='verifikasi').reduce((s,i)=>s+invTot(i).total,0);
  const aA=a=>'data-a="'+(a.sub?'sub':'go')+'" data-id="'+(a.sub||'')+'" data-v="'+(a.v||'antrean')+'"';
  const slaLate=subs.filter(s=>s.sentAt&&(s.status==='menunggu'||s.status==='review')&&slaLeft(s)<0).length;
  const top=att[0], done=subs.filter(s=>s.status==='disetujui').length;
  const fc=focusCard({eb:'FOKUS HARI INI',h:top?'Prioritas teratas':'Tidak ada yang mendesak',t:top?esc(top.t):'Semua pekerjaan terkendali.',p:top?esc(top.s):'Antrean review dan pengingat berjalan sesuai jadwal.',cta:top?'<button class="focus-time" '+aA(top)+'>'+ic('arrow')+'Buka sekarang<span>'+att.length+' perlu perhatian</span></button>':'',
    side:'<div class="focus-stats"><div class="'+(slaLate?'warn':'')+'"><b>'+slaLate+'</b><span>melewati SLA</span></div><div class="'+(late.length?'warn':'')+'"><b>'+late.length+'</b><span>klien terlambat</span></div></div>',f:[(rv.length+qc.length+q.length)+' laporan dalam proses',done+' dinyatakan sesuai']});
  const qa=quick([['data-a="go" data-v="antrean"','check','Antrean review','Periksa laporan yang masuk'],['data-a="go" data-v="klien"','users','Klien','Profil, layanan, dan dokumen'],isAdminDiv()?['data-a="go" data-v="invoice"','receipt','Invoice','Buat, kirim, dan verifikasi']:['data-a="go" data-v="pesan"','chat','Pesan','Percakapan dengan klien'],['data-a="go" data-v="pengingat"','clock','Pengingat','Atur dan kirim pengingat']]);
  const stats='<div class="grid g4 dashboard-stats">'+stat('white','inbox','Menunggu review',String(q.length),q.length?'Segera ditugaskan':'Tidak ada antrean','data-a="go" data-v="antrean"')+stat('blue','filecheck','Sedang direview',String(rv.length+qc.length),qc.length?qc.length+' menunggu rilis mutu':'Semua dalam SLA','data-a="go" data-v="antrean"')+stat(late.length?'peach':'white','hourglass','Klien terlambat kirim',String(late.length),late.length?'Pengingat perlu dikirim':'Semua tepat waktu','data-a="go" data-v="pengingat"')+(isAdminDiv()?stat('mint','wallet','Invoice terbuka',short(out),'Belum lunas','data-a="go" data-v="invoice"'):stat('mint','alert','Perlu perbaikan klien',String(perb.length),'Menunggu tanggapan klien','data-a="go" data-v="antrean"'))+'</div>';
  const per=Math.max(0,...subs.map(s=>s.t)), cur=subs.filter(s=>s.t===per), ks=['belum','menunggu','review','qc','perbaikan','disetujui'], tot=cur.length||1;
  const due=subs.filter(s=>s.status!=='disetujui').sort((a,b)=>dueOf(a)-dueOf(b)).slice(0,4);
  const stBlock='<section class="card card-pad"><div class="performance-head"><div><span class="eyebrow">STATUS LAPORAN · '+(per?periodLabel(per).toUpperCase():'')+'</span><h2>Sebaran laporan periode berjalan</h2></div><button class="link" data-a="go" data-v="antrean">Antrean review</button></div>'+
    '<div class="stack-bar" role="img" aria-label="Sebaran status laporan">'+ks.map(k=>{ const n=cur.filter(s=>s.status===k).length; return n?'<i style="width:'+(n/tot*100)+'%;background:'+STCOL[k]+'" title="'+STLAB[k]+': '+n+'"></i>':''; }).join('')+'</div>'+
    '<div class="stack-legend">'+ks.map(k=>'<span><span class="dot" style="background:'+STCOL[k]+'"></span>'+STLAB[k]+'<b>'+cur.filter(s=>s.status===k).length+'</b></span>').join('')+'</div>'+
    '<h3 class="sub-title" style="margin-top:20px">Tenggat terdekat</h3>'+(due.length?'<div class="deadline-list">'+due.map(s=>'<button class="deadline-row" data-a="sub" data-id="'+s.id+'"><div class="dl-date"><b>'+dueOf(s).getDate()+'</b><span>'+BLN3[dueOf(s).getMonth()]+'</span></div><div class="grow" style="min-width:0"><b class="trunc">'+esc(cl(s.cid).nama)+'</b><small>'+periodLabel(s.t)+'</small>'+bar(PROG[s.status],STCOL[s.status])+'</div>'+subPill(s,false)+'</button>').join('')+'</div>':'<div class="empty">Tidak ada tenggat terbuka.</div>')+'</section>';
  let right;
  if(isAdminDiv()){ const iv=INVOICES.filter(i=>i.status!=='draft').sort((a,b)=>a.due-b.due).slice(0,4), pc=i=>i.status==='lunas'?100:i.status==='verifikasi'?85:0;
    right='<section class="card card-pad"><div class="performance-head"><div><span class="eyebrow">PEMBAYARAN KLIEN</span><h2>Invoice terbaru</h2></div><button class="link" data-a="go" data-v="invoice">Semua</button></div><div class="pay-pills">'+pill('good',INVOICES.filter(i=>i.status==='lunas').length+' lunas')+pill('info',INVOICES.filter(i=>i.status==='verifikasi').length+' verifikasi')+pill('warn',INVOICES.filter(i=>i.status==='terkirim').length+' terkirim')+'</div><div class="pay-list">'+iv.map(i=>'<button class="pay-row" data-a="go" data-v="invoice"><div class="grow" style="min-width:0"><b class="trunc">'+esc(cl(i.cid).nama)+'</b><small>'+i.no+' · jatuh tempo '+dfmt(i.due)+'</small></div><div class="pay-meter">'+bar(pc(i),pc(i)===100?'var(--good)':'')+'<span>'+short(invTot(i).total)+' '+invStatus(i)+'</span></div></button>').join('')+'</div></section>'; }
  else { const fl=[...new Set(subs.map(s=>s.cid))].map(id=>({c:cl(id),n:subs.filter(s=>s.cid===id).reduce((a,s)=>a+s.findings.filter(f=>f.st==='terbuka').length,0)})).sort((a,b)=>b.n-a.n).slice(0,5);
    right='<section class="card card-pad"><div class="performance-head"><div><span class="eyebrow">TEMUAN TERBUKA</span><h2>Per klien</h2></div></div><div class="load-list">'+fl.map(x=>'<div class="load-row"><span>'+esc(x.c.singkat||x.c.nama)+'</span>'+bar(Math.min(100,x.n*20))+'<b>'+x.n+'</b></div>').join('')+'</div></section>'; }
  let load='';
  if(su){ const tim=USERS.filter(x=>x.role==='tim'||x.role==='mitra'); const wl=x=>ENG.filter(e=>e.pic===x.id||e.mitra===x.id).length+SUBS.filter(s=>['menunggu','review','qc'].includes(s.status)&&(reviewerOf(s)===x.id||mitraOf(s)===x.id)).length; const mx=Math.max(1,...tim.map(wl));
    load='<section class="card card-pad"><span class="eyebrow">BEBAN KERJA</span><h2 class="sub-title">Tim dan mitra</h2><div class="load-list">'+tim.map(x=>'<div class="load-row"><span>'+esc(x.nama)+'</span>'+bar(Math.round(wl(x)/mx*100))+'<b>'+wl(x)+'</b></div>').join('')+'</div></section>'; }
  const attC='<section class="card card-pad"><div class="performance-head"><div><span class="eyebrow">PERLU PERHATIAN</span><h2>Daftar tindak lanjut</h2></div>'+pill('mute',String(att.length))+'</div>'+(att.length?'<div class="list" style="margin-top:12px">'+att.slice(0,7).map(a=>'<button class="item itemlink stripe '+a.sev+'" '+aA(a)+'><div class="grow"><strong>'+esc(a.t)+'</strong><small>'+esc(a.s)+'</small></div></button>').join('')+'</div>':'<div class="empty">Tidak ada yang perlu perhatian.</div>')+'</section>';
  const fin=su?(()=>{ const L=finLatest(CLIENTS.map(c=>c.id)); return '<div class="performance-head" style="margin:-4px 0 -6px"><div><span class="eyebrow">KEUANGAN KLIEN · GABUNGAN</span></div><button class="link" data-a="go" data-v="keuangan">Buka laporan keuangan</button></div>'+finStrip([['white','trend','Pendapatan klien',short(L.pend),'Periode terkini tiap klien','data-a="go" data-v="keuangan"'],['mint','wallet','Laba bersih',short(L.laba),L.pend?(L.laba/L.pend*100).toFixed(1).replace('.',',')+'% dari pendapatan':''],['blue','wallet','Saldo kas akhir',short(L.kas),'Total seluruh klien'],['white','receipt','Piutang usaha klien',short(L.ar),'Invoice usaha belum dibayar']]); })():'';
  return intro(su?'Gambaran menyeluruh: laporan klien, tim, mitra, dan penagihan.':'Ringkasan pekerjaan yang ditugaskan kepada Anda.')+'<div class="workspace-overview">'+fc+qa+'</div>'+stats+fin+
    '<div class="dashboard-detail-grid">'+stBlock+right+'</div><div class="dashboard-detail-grid">'+attC+'<div class="col gap-4">'+load+feed(actFeed(subs,false))+'</div></div>'+foot();
}
function dashMitra(){
  const u=me(), js=mySubs(), act=js.filter(s=>['menunggu','review'].includes(s.status)), done=js.filter(s=>s.status==='qc'||s.status==='disetujui'||s.status==='perbaikan');
  const nx=act.sort((a,b)=>slaLeft(a)-slaLeft(b))[0];
  const fc=focusCard({eb:'FOKUS HARI INI',h:nx?'Review terdekat Anda':'Tidak ada review aktif',t:nx?esc(cl(nx.cid).nama)+' · '+periodLabel(nx.t):'Semua pekerjaan beres.',p:nx?(slaLeft(nx)<0?'Melewati SLA '+(-slaLeft(nx))+' hari kerja':'Sisa SLA '+slaLeft(nx)+' hari kerja')+' · dikirim '+dfmt(nx.sentAt):'Pekerjaan baru akan muncul saat ditugaskan.',cta:nx?'<button class="focus-time" data-a="sub" data-id="'+nx.id+'">'+ic('arrow')+'Buka review<span>'+act.length+' aktif</span></button>':'',f:[done.length+' selesai periode ini','Honor '+rp(u.fee)+' per review']});
  return intro('Review yang ditugaskan kepada Anda oleh tim '+CONFIG.brand+'.')+'<div class="workspace-overview">'+fc+quick([['data-a="go" data-v="antrean"','check','Daftar review','Seluruh pekerjaan Anda'],['data-a="go" data-v="notif"','bell','Notifikasi','Penugasan dan pembaruan']])+'</div>'+
  '<div class="grid g3 dashboard-stats">'+stat('white','inbox','Review aktif',String(act.length),'Perlu dikerjakan','data-a="go" data-v="antrean"')+stat('mint','filecheck','Selesai periode ini',String(done.length),'Termasuk menunggu rilis','data-a="go" data-v="antrean"')+stat('blue','wallet','Honor per review',rp(u.fee),'Total berjalan '+rp(u.fee*done.length),'')+'</div>'+
  '<div class="dashboard-detail-grid"><section class="card card-pad"><div class="performance-head"><div><span class="eyebrow">DAFTAR PEKERJAAN</span><h2>Review yang ditugaskan</h2></div></div>'+(js.length?'<div class="list" style="margin-top:12px">'+js.map(s=>'<button class="item itemlink" data-a="sub" data-id="'+s.id+'"><div class="grow"><strong>'+esc(cl(s.cid).nama)+' · '+periodLabel(s.t)+'</strong><small>'+(s.sentAt?'Dikirim '+dfmt(s.sentAt)+' · sisa SLA '+slaLeft(s)+' hari kerja':'Belum dikirim klien')+'</small></div>'+subPill(s,false)+'</button>').join('')+'</div>':'<div class="empty">Belum ada pekerjaan.</div>')+'</section>'+feed(actFeed(js,false))+'</div>'+foot();
}



function canFin(cid){ const u=me(); if(!cl(cid)) return false; if(u.role==='superuser') return true; if(u.role==='klien') return u.cid===cid; return (cl(cid).finAccess||[]).includes(u.id); }
const finClients=()=>CLIENTS.filter(c=>canFin(c.id));
const logoBox=(c,n)=>{ const b=c.brand; if(b.logo) return '<img class="lg-img" src="'+b.logo+'" alt="Logo '+esc(c.nama)+'" style="width:'+n+'px;height:'+n+'px">'; const w=c.nama.replace(/^(PT|CV|UD)\s+/i,'').split(/\s+/).slice(0,2).map(x=>x[0]).join('').toUpperCase(); return '<div class="lg-mono" style="width:'+n+'px;height:'+n+'px;background:'+(b.warna||'#7A2418')+';font-size:'+Math.round(n*.38)+'px" aria-hidden="true">'+w+'</div>'; };
const cinvStatus=i=>{ if(i.status==='lunas') return pill('good','Lunas'); if(i.status==='draft') return pill('mute','Draft'); const d=daysTo(i.due); return d<0?pill('bad','Terlambat '+(-d)+' hari'):pill('warn','Jatuh tempo '+dfmt(i.due)); };
const locked=(t,txt,btn)=>'<div class="locked"><div>'+ic('lock')+'<h2>'+t+'</h2><p>'+txt+'</p>'+(btn||'')+'</div></div>';
function cinvHtml(i,plain){ const c=cl(i.cid), b=c.brand, col=b.warna||'#7A2418';
  return invDoc({i,col,logo:logoBox(c,54),brand:'<div class="ient" style="font-size:16px">'+esc(c.nama)+'</div>',lines:[b.alamat,[b.email,b.telp].filter(Boolean).join(' · '),'NPWP '+c.npwp],to:{n:i.cust||'-',l:[i.addr,i.email]},bank:b.rek||'-',note:(i.note?i.note+' ':'')+(b.catatan||''),status:plain?'':cinvStatus(i)}); }
function cinvText(i){ const t=invTot(i), c=cl(i.cid); return 'INVOICE '+i.no+'\n'+c.nama+'\nKepada: '+i.cust+'\nJatuh tempo: '+dfmt(i.due)+'\n\n'+i.lines.map(l=>'- '+l.desc+' ('+l.qty+' x '+rp(l.harga)+')').join('\n')+'\n\nSubtotal: '+rp(t.dpp)+(i.ppn?'\nPPN 11%: '+rp(t.ppn):'')+'\nTotal: '+rp(t.total)+'\n'+c.brand.rek; }
function bars(labels,series){ const mx=Math.max(1,...series.flatMap(x=>x.vals)); return '<div class="bars" role="img" aria-label="Grafik '+series.map(x=>x.name).join(' dan ')+' per periode">'+labels.map((l,i)=>'<div class="bg"><div class="bs">'+series.map(x=>'<i style="height:'+Math.max(2,x.vals[i]/mx*100)+'%;background:'+x.color+'" title="'+esc(x.name)+': '+short(x.vals[i])+'"></i>').join('')+'</div><span>'+l+'</span></div>').join('')+'</div><div class="stack-legend">'+series.map(x=>'<span><span class="dot" style="background:'+x.color+'"></span>'+x.name+'</span>').join('')+'</div>'; }
function stmt(kind,f){ const r=(l,v,b)=>'<tr><td>'+(b?'<strong>':'')+l+(b?'</strong>':'')+'</td><td class="r num">'+(b?'<strong>':'')+rp(v)+(b?'</strong>':'')+'</td></tr>'; let rows='';
  if(kind==='lr') rows=r('Total pendapatan',f.pendapatan)+r('Total beban (termasuk pajak)',f.beban)+r('Laba bersih',f.laba,true)+'<tr><td>Margin laba bersih</td><td class="r num">'+(f.pendapatan?(f.laba/f.pendapatan*100).toFixed(1).replace('.',','):'0')+'%</td></tr>';
  else if(kind==='nr'){ const d=f.aset-(f.liab+f.ekuitas); rows=r('Total aset',f.aset,true)+r('Total liabilitas',f.liab)+r('Total ekuitas',f.ekuitas)+'<tr><td>Selisih aset dengan liabilitas + ekuitas</td><td class="r num">'+(d===0?pill('good','Seimbang'):pill('bad',rp(d)))+'</td></tr>'; }
  else rows=r('Saldo kas awal',f.kas0)+r('Total penerimaan kas',f.masuk)+r('Total pengeluaran kas',f.keluar)+r('Saldo kas akhir',f.kas1,true);
  return '<div class="tbl-wrap"><table><tbody>'+rows+'</tbody></table></div>'; }
function finStrip(rows){ return '<div class="grid g4 dashboard-stats">'+rows.map(x=>stat(x[0],x[1],x[2],x[3],x[4],x[5]||'')).join('')+'</div>'; }
function finLatest(cids){ let pend=0,laba=0,kas=0,ar=0; cids.forEach(id=>{ const fs=figSubs(id), l=fs[fs.length-1]; if(l){ pend+=l.fig.pendapatan; laba+=l.fig.laba; kas+=l.fig.kas1; } ar+=arOf(id); }); return {pend,laba,kas,ar}; }


function vCinv(){
  const c=cl(me().cid), f=S.cinvF, all=CINV.filter(i=>i.cid===c.id), rows=all.filter(i=>f==='semua'||(f==='terbuka'?i.status==='terkirim':i.status===f)).sort((a,b)=>b.tgl-a.tgl||b.no.localeCompare(a.no));
  const ar=arOf(c.id), late=all.filter(i=>i.status==='terkirim'&&daysTo(i.due)<0), paid=all.filter(i=>i.status==='lunas').reduce((a,i)=>a+invTot(i).total,0);
  return head('Invoice usaha','Buat dan kirim invoice kepada pelanggan Anda dengan logo dan identitas usaha sendiri.','<button class="btn" data-a="go" data-v="usaha">'+ic('gear')+'Profil usaha</button><button class="btn primary" data-a="cinvnew">'+ic('plus')+'Buat invoice</button>')+
  finStrip([['white','wallet','Piutang berjalan',short(ar),all.filter(i=>i.status==='terkirim').length+' invoice belum dibayar','data-a="cinvf" data-k="terbuka"'],[late.length?'peach':'white','alert','Terlambat dibayar',String(late.length),late.length?short(late.reduce((a,i)=>a+invTot(i).total,0))+' menunggu pelanggan':'Tidak ada','data-a="cinvf" data-k="terbuka"'],['mint','filecheck','Sudah dibayar',short(paid),'Seluruh periode','data-a="cinvf" data-k="lunas"'],['blue','receipt','Draft',String(all.filter(i=>i.status==='draft').length),'Belum dikirim','data-a="cinvf" data-k="draft"']])+
  '<div class="card"><div class="card-h"><div class="filters">'+[['semua','Semua'],['draft','Draft'],['terbuka','Belum dibayar'],['lunas','Lunas']].map(([k,l])=>'<button data-a="cinvf" data-k="'+k+'" aria-pressed="'+(f===k)+'">'+l+'</button>').join('')+'</div></div>'+
  (rows.length?'<div class="tbl-wrap"><table><thead><tr><th>No. invoice</th><th>Pelanggan</th><th>Tanggal</th><th>Jatuh tempo</th><th class="r">Total</th><th>Status</th></tr></thead><tbody>'+rows.map(i=>'<tr class="click" data-a="cinvopen" data-no="'+i.no+'" tabindex="0"><td class="mono">'+i.no+jenisTag(i)+'</td><td>'+esc(i.cust)+'</td><td class="num">'+dfmt(i.tgl)+'</td><td class="num">'+dfmt(i.due)+'</td><td class="r num"><strong>'+num(invTot(i).total)+'</strong></td><td>'+cinvStatus(i)+'</td></tr>').join('')+'</tbody></table></div>':'<div class="empty">Belum ada invoice pada filter ini.</div>')+'</div>'+
  '<p class="note">Invoice ini dari usaha Anda untuk pelanggan Anda. Berbeda dengan menu Tagihan, yaitu invoice jasa '+CONFIG.brand+' kepada Anda. Data ini dapat dilihat Managing Director dan tim yang diberi akses.</p>';
}
/* ---------- klien: profil usaha ---------- */
function kopHtml(c){ const b=c.brand; return '<div class="kop"><div style="display:flex;gap:12px;align-items:center">'+logoBox(c,54)+'<div><strong>'+esc(c.nama)+'</strong><div class="note">'+esc(b.alamat)+'<br>'+esc(b.email)+' · '+esc(b.telp)+'</div></div></div><div class="kop-line" style="background:'+(b.warna||'#7A2418')+'"></div><div style="display:flex;justify-content:space-between;align-items:baseline"><b style="color:'+(b.warna||'#7A2418')+';font:700 18px var(--f-display)">INVOICE</b><span class="mono">'+esc(b.pre)+'-2611-001</span></div></div>'; }
function accessList(c){ const names=USERS.filter(x=>(c.finAccess||[]).includes(x.id)); return '<ul class="acc-list"><li>'+ic('shield')+'<span><b>Manda</b><small>Managing Director · selalu memiliki akses</small></span></li>'+names.map(x=>'<li>'+ic('eye')+'<span><b>'+esc(x.nama)+'</b><small>'+esc(x.role==='tim'?DIV[x.div]:x.jabatan)+'</small></span></li>').join('')+'</ul>'; }
function vUsaha(){
  const c=cl(me().cid), b=c.brand;
  const sw=['#7A2418','#B04A2C','#2F5D73','#1B7A4B','#5B3A8C','#2B2B2B'].map(x=>'<button class="swatch" data-a="profcolor" data-c="'+x+'" style="background:'+x+'" aria-label="Warna '+x+'" aria-pressed="'+((b.warna||'')===x)+'"></button>').join('');
  const inp=(l,k,ph,full)=>'<label class="f" style="'+(full?'grid-column:1/-1':'')+'">'+l+'<input type="text" data-input="prof" data-k="'+k+'" value="'+esc(b[k]||'')+'" placeholder="'+(ph||'')+'"></label>';
  return head('Profil usaha','Identitas ini tampil pada setiap invoice yang Anda terbitkan kepada pelanggan.','<button class="btn primary" data-a="profsave">Simpan profil</button>')+
  '<div class="dashboard-detail-grid"><div class="col gap-4"><section class="card card-pad"><span class="eyebrow">LOGO USAHA</span><div class="logo-up"><div id="logoPrev">'+logoBox(c,88)+'</div><div class="col" style="gap:8px"><label class="btn" style="cursor:pointer">'+ic('upload')+'Unggah logo<input type="file" accept="image/png,image/jpeg,image/svg+xml,image/webp" data-change="logoup" hidden></label>'+(b.logo?'<button class="btn sm" data-a="logorm">Hapus logo</button>':'')+'<small class="note">PNG, JPG, SVG, atau WebP, maksimal 600 KB. Tanpa logo, inisial usaha dipakai.</small></div></div></section>'+
  '<section class="card card-pad"><span class="eyebrow">DATA USAHA</span><div class="formgrid"><label class="f" style="grid-column:1/-1">Nama usaha<input type="text" value="'+esc(c.nama)+'" disabled></label>'+inp('Alamat','alamat','',true)+inp('Email','email')+inp('Telepon','telp')+inp('Rekening pembayaran','rek','Bank, nomor, atas nama',true)+inp('Awalan nomor invoice','pre','MBN')+inp('Catatan kaki invoice','catatan','',false)+'<div style="grid-column:1/-1"><div class="note" style="margin-bottom:6px">Warna aksen invoice</div><div class="swatches">'+sw+'</div></div></div><p class="note">NPWP '+esc(c.npwp)+' diubah lewat tim '+CONFIG.brand+'.</p></section></div>'+
  '<div class="col gap-4"><section class="card card-pad"><span class="eyebrow">PRATINJAU KOP INVOICE</span><div id="kopPrev" style="margin-top:12px">'+kopHtml(c)+'</div></section>'+
  '<section class="card card-pad"><span class="eyebrow">AKSES DATA KEUANGAN</span><h2 class="sub-title" style="margin-top:6px">Siapa yang dapat melihat laporan keuangan Anda</h2>'+accessList(c)+'<p class="note">Hak akses diatur oleh Managing Director '+CONFIG.brand+'. Setiap perubahan akan diberitahukan kepada Anda.</p></section></div></div>';
}
/* ---------- staf: keuangan klien ---------- */
function vKeuangan(){
  const u=me(), su=isSU(), cs=finClients();
  if(!su) S.ftab='lapor';
  if(S.fkc==='all'&&!su) S.fkc=null;
  if(!S.fkc||(S.fkc!=='all'&&!CLIENTS.some(c=>c.id===S.fkc))) S.fkc=su?'all':(cs[0]&&cs[0].id);
  const chips='<div class="filters">'+(su?'<button data-a="fkc" data-k="all" aria-pressed="'+(S.fkc==='all')+'">Gabungan</button>':'')+CLIENTS.filter(c=>su||canFin(c.id)).map(c=>'<button data-a="fkc" data-k="'+c.id+'" aria-pressed="'+(S.fkc===c.id)+'">'+(canFin(c.id)?'':ic('lock')+' ')+esc(c.singkat||c.nama)+'</button>').join('')+'</div>';
  const tabs=su?'<div class="tabs" role="tablist">'+[['lapor','Laporan keuangan'],['akses','Hak akses']].map(([k,l])=>'<button role="tab" aria-selected="'+(S.ftab===k)+'" data-a="ftab" data-k="'+k+'">'+l+'</button>').join('')+'</div>':'';
  const top=head('Keuangan klien',su?'Laporan keuangan seluruh klien. Anda dapat mengatur siapa yang boleh melihat data keuangan tiap klien.':'Laporan keuangan klien yang diberikan aksesnya kepada Anda.')+'<div class="row" style="justify-content:space-between;align-items:center">'+chips+tabs+'</div>';
  if(su&&S.ftab==='akses') return top+vFinAkses();
  if(S.fkc==='all') return top+vFinGabungan();
  return top+vFinKlien(cl(S.fkc));
}
function vFinGabungan(){
  const L=finLatest(CLIENTS.map(c=>c.id)), per=[8,9,10];
  const sum=(t,k)=>CLIENTS.reduce((a,c)=>{ const s=figSubs(c.id).find(x=>x.t===t); return a+(s?s.fig[k]:0); },0);
  const rows=CLIENTS.map(c=>{ const fs=figSubs(c.id), l=fs[fs.length-1]; if(!l) return ''; const f=l.fig; return '<tr class="click" data-a="fkc" data-k="'+c.id+'" tabindex="0"><td><strong>'+esc(c.nama)+'</strong></td><td>'+periodLabel(l.t)+'</td><td class="r num">'+num(f.pendapatan)+'</td><td class="r num">'+num(f.beban)+'</td><td class="r num"><strong>'+num(f.laba)+'</strong></td><td class="r num">'+(f.laba/f.pendapatan*100).toFixed(1).replace('.',',')+'%</td><td class="r num">'+num(f.kas1)+'</td><td class="r num">'+num(arOf(c.id))+'</td><td>'+subPill(l,false)+'</td></tr>'; }).join('');
  return finStrip([['white','trend','Pendapatan klien',short(L.pend),'Periode terkini tiap klien'],['mint','wallet','Laba bersih',short(L.laba),(L.pend?(L.laba/L.pend*100).toFixed(1).replace('.',','):'0')+'% dari pendapatan'],['blue','wallet','Saldo kas akhir',short(L.kas),'Total seluruh klien'],['white','receipt','Piutang usaha klien',short(L.ar),'Invoice usaha belum dibayar']])+
  '<div class="dashboard-detail-grid"><section class="card card-pad"><div class="performance-head"><div><span class="eyebrow">TREN GABUNGAN</span><h2>Pendapatan dan laba bersih</h2></div></div>'+bars(per.map(t=>BLN3[t-1]),[{name:'Pendapatan',color:'var(--brand)',vals:per.map(t=>sum(t,'pendapatan'))},{name:'Laba bersih',color:'var(--copper)',vals:per.map(t=>sum(t,'laba'))}])+'</section>'+
  '<section class="card card-pad"><span class="eyebrow">AKSES</span><h2 class="sub-title" style="margin-top:6px">Siapa melihat apa</h2><div class="load-list">'+CLIENTS.map(c=>'<div class="load-row" style="grid-template-columns:minmax(0,1fr) auto"><span>'+esc(c.singkat||c.nama)+'</span><b>'+(c.finAccess.length+1)+' orang</b></div>').join('')+'</div><button class="link" style="margin-top:10px" data-a="ftab" data-k="akses">Atur hak akses</button></section></div>'+
  '<div class="card"><div class="card-h"><h2>Ringkasan per klien</h2><span class="grow"></span><span class="note">Angka dilaporkan klien, status menunjukkan hasil review</span></div><div class="tbl-wrap"><table><thead><tr><th>Klien</th><th>Periode</th><th class="r">Pendapatan</th><th class="r">Beban</th><th class="r">Laba bersih</th><th class="r">Margin</th><th class="r">Kas akhir</th><th class="r">Piutang usaha</th><th>Review</th></tr></thead><tbody>'+rows+'</tbody></table></div></div>';
}
function vFinKlien(c){
  if(!canFin(c.id)) return locked('Akses data keuangan belum diberikan','Anda belum memiliki izin untuk melihat laporan keuangan '+esc(c.nama)+'. Managing Director dapat memberikan akses.','<button class="btn primary" data-a="finreq" data-cid="'+c.id+'">Minta akses</button>');
  const fs=figSubs(c.id); if(!fs.length) return '<div class="card"><div class="empty">Belum ada laporan keuangan yang dikirim klien ini.</div></div>';
  if(!S.fkp||!fs.some(x=>x.t===S.fkp)||S.fkpc!==c.id){ S.fkp=fs[fs.length-1].t; S.fkpc=c.id; }
  const cur=fs.find(x=>x.t===S.fkp), f=cur.fig, kind=S.fst||'lr';
  const open=cur.findings.filter(x=>x.st==='terbuka').length, inv=CINV.filter(i=>i.cid===c.id).sort((a,b)=>b.tgl-a.tgl).slice(0,5);
  return '<div class="filters">'+fs.map(x=>'<button data-a="fkp" data-k="'+x.t+'" aria-pressed="'+(S.fkp===x.t)+'">'+periodLabel(x.t)+'</button>').join('')+'</div>'+
  (cur.status==='perbaikan'?'<div class="banner warn"><div><strong>Laporan '+periodLabel(cur.t)+' masih dalam perbaikan.</strong>Angka dapat berubah setelah klien mengirim ulang ('+open+' temuan terbuka).</div></div>':cur.status==='disetujui'?'<div class="banner good"><div><strong>Laporan '+periodLabel(cur.t)+' dinyatakan sesuai oleh reviewer.</strong>Angka di bawah sudah melewati pemeriksaan.</div></div>':'<div class="banner info"><div><strong>Laporan '+periodLabel(cur.t)+' belum selesai direview.</strong>Angka masih sebagaimana dilaporkan klien.</div></div>')+
  finStrip([['white','trend','Pendapatan',short(f.pendapatan),periodLabel(cur.t)],['mint','wallet','Laba bersih',short(f.laba),(f.laba/f.pendapatan*100).toFixed(1).replace('.',',')+'% dari pendapatan'],['blue','wallet','Saldo kas akhir',short(f.kas1),'Awal '+short(f.kas0)],['white','receipt','Piutang usaha klien',short(arOf(c.id)),'Dari invoice usaha klien']])+
  '<div class="dashboard-detail-grid"><section class="card card-pad"><div class="performance-head"><div><span class="eyebrow">LAPORAN KEUANGAN · '+periodLabel(cur.t).toUpperCase()+'</span><h2>'+esc(c.nama)+'</h2></div><div class="tabs" role="tablist">'+[['lr','Laba rugi'],['nr','Neraca'],['ak','Arus kas']].map(([k,l])=>'<button role="tab" aria-selected="'+(kind===k)+'" data-a="fst" data-k="'+k+'">'+l+'</button>').join('')+'</div></div><div style="margin-top:12px">'+stmt(kind,f)+'</div></section>'+
  '<section class="card card-pad"><div class="performance-head"><div><span class="eyebrow">TREN</span><h2>Pendapatan dan laba</h2></div></div>'+bars(fs.map(x=>BLN3[x.t-1]),[{name:'Pendapatan',color:'var(--brand)',vals:fs.map(x=>x.fig.pendapatan)},{name:'Laba bersih',color:'var(--copper)',vals:fs.map(x=>x.fig.laba)}])+'<button class="link" style="margin-top:12px" data-a="sub" data-id="'+cur.id+'">Buka detail review '+periodLabel(cur.t)+'</button></section></div>'+
  '<div class="card"><div class="card-h"><h2>Invoice usaha klien</h2><span class="grow"></span><span class="note">Hanya baca</span></div>'+(inv.length?'<div class="tbl-wrap"><table><thead><tr><th>No. invoice</th><th>Pelanggan</th><th>Jatuh tempo</th><th class="r">Total</th><th>Status</th></tr></thead><tbody>'+inv.map(i=>'<tr><td class="mono">'+i.no+jenisTag(i)+'</td><td>'+esc(i.cust)+'</td><td class="num">'+dfmt(i.due)+'</td><td class="r num"><strong>'+num(invTot(i).total)+'</strong></td><td>'+cinvStatus(i)+'</td></tr>').join('')+'</tbody></table></div>':'<div class="empty">Belum ada invoice usaha.</div>')+'</div>';
}
function vFinAkses(){
  const staff=USERS.filter(x=>x.role==='tim'||x.role==='mitra');
  return '<div class="card"><div class="card-h"><h2>Hak akses data keuangan klien</h2><span class="grow"></span><span class="note">Managing Director selalu memiliki akses</span></div><div class="tbl-wrap"><table class="matrix"><thead><tr><th>Tim dan mitra</th>'+CLIENTS.map(c=>'<th>'+esc(c.singkat||c.nama)+'</th>').join('')+'</tr></thead><tbody>'+
  '<tr><td><strong>Manda</strong><div class="note">Managing Director</div></td>'+CLIENTS.map(()=>'<td><input type="checkbox" checked disabled aria-label="Selalu aktif"></td>').join('')+'</tr>'+
  staff.map(x=>'<tr><td><strong>'+esc(x.nama)+'</strong><div class="note">'+esc(x.role==='tim'?DIV[x.div]:x.jabatan)+'</div></td>'+CLIENTS.map(c=>'<td><input type="checkbox" data-change="finacc" data-cid="'+c.id+'" data-uid="'+x.id+'"'+(c.finAccess.includes(x.id)?' checked':'')+' aria-label="'+esc(x.nama)+' untuk '+esc(c.nama)+'"></td>').join('')+'</tr>').join('')+'</tbody></table></div></div>'+
  '<div class="card card-pad"><span class="eyebrow">LOG PERUBAHAN AKSES</span><div class="activity-timeline">'+FINLOG.slice(0,8).map(a=>'<div class="activity-item"><span class="activity-icon">'+ic('shield')+'</span><div><b>'+esc(a.text)+'</b><small>'+esc(a.who)+' · '+esc(a.ts)+'</small></div></div>').join('')+'</div></div>'+
  '<p class="note">Klien melihat daftar orang yang memiliki akses pada halaman Profil usaha dan menerima notifikasi saat akses berubah.</p>';
}

/* ---------- kirim laporan (klien) ---------- */
function openForm(s){ S.form={sid:s.id,f:Object.fromEntries(FIG.map(([k])=>[k,s.fig?s.fig[k]:''])),files:{...s.files},catatan:s.catatan||''}; }
function vKirim(){
  const u=me(), c=cl(u.cid), subs=SUBS.filter(s=>s.cid===c.id).sort((a,b)=>b.t-a.t);
  const open=subs.filter(s=>s.status==='belum'||s.status==='perbaikan');
  if(!open.length) return head('Kirim laporan keuangan','Unggah laporan periode berjalan untuk direview tim.')+'<div class="card empty">Semua periode sudah dikirim. Periode '+periodLabel(11)+' dibuka 1 Desember 2026.</div>'+hasilList(subs);
  if(!S.form||!open.some(s=>s.id===S.form.sid)) openForm(open[0]);
  const s=SUBS.find(x=>x.id===S.form.sid), f=S.form;
  const fields=FIG.map(([k,l])=>'<label class="f">'+l+' (Rp)<input type="number" step="1" inputmode="numeric" data-input="fld" data-k="'+k+'" value="'+esc(f.f[k])+'"></label>').join('');
  const files=DOCREQ.map(d=>'<div class="item"><div class="grow"><strong>'+d.nama+'</strong><small id="fn_'+d.k+'">'+(f.files[d.k]?esc(f.files[d.k]):'Belum diunggah')+'</small></div><label class="btn sm" style="cursor:pointer">'+ic('upload')+(f.files[d.k]?'Ganti':'Pilih berkas')+'<input type="file" data-change="formfile" data-k="'+d.k+'" hidden></label></div>').join('');
  const dl=daysTo(dueOf(s));
  return head('Kirim laporan keuangan','Isi ringkasan angka dan unggah dokumen. Sistem langsung memeriksa kelengkapan dan kewajaran sebelum dikirim ke tim.',
    open.length>1?'<select id="perSel" data-change="formper" aria-label="Periode">'+open.map(x=>'<option value="'+x.id+'"'+(x.id===s.id?' selected':'')+'>'+periodLabel(x.t)+(x.status==='perbaikan'?' (perbaikan)':'')+'</option>').join('')+'</select>':'')+
  '<div class="banner '+(s.status==='perbaikan'?'bad':dl<0?'bad':'info')+'"><div><strong>'+periodLabel(s.t)+(s.status==='perbaikan'?' · pengiriman ulang':'')+'</strong>'+(s.status==='perbaikan'?'Perbaiki angka atau dokumen sesuai temuan, lalu kirim ulang sebelum '+dfmt(dueOf(s))+'.':dl<0?'Batas kirim '+dfmt(dueOf(s))+' sudah lewat '+(-dl)+' hari. Kirim secepatnya agar review tidak tertunda.':'Batas kirim '+dfmt(dueOf(s))+'.')+'</div></div>'+
  '<div class="grid g2"><div style="display:flex;flex-direction:column;gap:16px"><div class="card"><div class="card-h"><h2>1. Angka ringkasan laporan</h2><span class="grow"></span><button class="btn sm" data-a="fillfig">Isi otomatis dari buku besar</button></div><div class="card-b"><div class="grid g2e">'+fields+'</div></div></div>'+
   '<div class="card"><div class="card-h"><h2>2. Dokumen wajib</h2></div><div class="list">'+files+'</div></div>'+
   '<div class="card"><div class="card-h"><h2>3. Catatan untuk reviewer</h2></div><div class="card-b"><textarea id="catatan" data-input="catatan" placeholder="Jelaskan transaksi tidak biasa, perubahan besar, atau ketidaksesuaian angka (wajib bila pra-periksa gagal).">'+esc(f.catatan)+'</textarea></div></div></div>'+
   '<div style="display:flex;flex-direction:column;gap:16px"><div class="card" style="position:sticky;top:0"><div class="card-h"><h2>Pra-periksa otomatis</h2><span class="grow"></span><span id="chkSum"></span></div><div id="chk"></div><div class="card-b" style="border-top:1px solid var(--line);display:flex;gap:8px;flex-wrap:wrap"><button class="btn" data-a="formsave">Simpan draft</button><button class="btn primary" data-a="formsubmit">'+(s.status==='perbaikan'?'Kirim ulang untuk review':'Kirim untuk review')+'</button></div></div></div></div>'+hasilListCard(subs);
}
function chkHtml(){
  const f=S.form, s=SUBS.find(x=>x.id===f.sid), c=runChecks(f.f,f.files,prevFig(s.cid,s.t),f.catatan);
  const sum=checkSummary(c);
  return {html:c.map(x=>'<div class="ck '+x.lvl+'"><i>'+(x.lvl==='ok'?'✓':x.lvl==='warn'?'!':'×')+'</i><div><strong>'+x.t+'</strong><small>'+esc(x.d)+'</small></div></div>').join(''),sum:sum==='ok'?pill('good','Lolos'):sum==='warn'?pill('warn','Ada peringatan'):pill('bad','Perlu diperbaiki'),checks:c};
}
function hasilListCard(subs){ return '<div class="card"><div class="card-h"><h2>Status periode</h2></div><div class="list">'+subs.map(s=>'<button class="item itemlink" data-a="sub" data-id="'+s.id+'"><div class="grow"><strong>'+periodLabel(s.t)+'</strong><small>'+(s.sentAt?'Dikirim '+dfmt(s.sentAt):'Belum dikirim')+'</small></div>'+subPill(s,true)+'</button>').join('')+'</div></div>'; }
function hasilList(subs){ return hasilListCard(subs); }
function vHasil(){
  const subs=SUBS.filter(s=>s.cid===me().cid&&s.status!=='belum').sort((a,b)=>b.t-a.t);
  return head('Hasil review','Hasil pemeriksaan tim atas laporan keuangan Anda, beserta temuan dan tindak lanjutnya.')+
  '<div class="card"><div class="tbl-wrap"><table><thead><tr><th>Periode</th><th>Dikirim</th><th>Status</th><th class="r">Temuan terbuka</th><th></th></tr></thead><tbody>'+
  subs.map(s=>{ const o=s.findings.filter(f=>f.st!=='ditutup').length; const vis=['perbaikan','disetujui'].includes(s.status); return '<tr class="click" data-a="sub" data-id="'+s.id+'" tabindex="0"><td><strong>'+periodLabel(s.t)+'</strong></td><td class="num">'+dfmt(s.sentAt)+'</td><td>'+subPill(s,true)+'</td><td class="r num">'+(vis?o:'–')+'</td><td class="r"><span class="btn sm">Buka</span></td></tr>'; }).join('')+'</tbody></table></div></div>';
}

/* ---------- detail laporan (semua peran) ---------- */
function figTable(s){
  if(!s.fig) return '<div class="empty">Belum ada angka yang dilaporkan.</div>';
  const p=prevFig(s.cid,s.t);
  return '<div class="tbl-wrap"><table><thead><tr><th>Uraian</th><th class="r">'+periodLabel(s.t)+'</th>'+(p?'<th class="r">Periode lalu</th>':'')+'</tr></thead><tbody>'+FIG.map(([k,l])=>'<tr><td>'+l+'</td><td class="r num">'+num(s.fig[k])+'</td>'+(p?'<td class="r num" style="color:var(--muted)">'+num(p[k])+'</td>':'')+'</tr>').join('')+'</tbody></table></div>';
}
function vDetail(s){
  const u=me(), c=cl(s.cid), client=isKlien(), rv=!client&&canReview(s), mitra=isMitra();
  const checks=s.fig?runChecks(s.fig,s.files,prevFig(s.cid,s.t),s.catatan):[];
  const showF=!client||['perbaikan','disetujui'].includes(s.status);
  const finds=s.findings.filter(f=>showF||!client);
  let banner='';
  if(client){
    banner=s.status==='disetujui'?'<div class="banner good"><div><strong>Laporan dinyatakan sesuai.</strong>Direview oleh tim Mandagroup pada '+dfmt(s.doneAt)+'. Tidak ada tindakan lanjutan.</div></div>'
     :s.status==='perbaikan'?'<div class="banner bad"><div><strong>Laporan perlu diperbaiki.</strong>Tanggapi setiap temuan terbuka, perbaiki angka atau dokumen bila perlu, lalu kirim ulang sebelum '+dfmt(dueOf(s))+'.</div><div class="grow"></div><button class="btn primary" data-a="reopen" data-id="'+s.id+'">Perbarui dan kirim ulang</button></div>'
     :s.status==='belum'?'<div class="banner warn"><div><strong>Laporan belum dikirim.</strong>Lengkapi formulir pengiriman laporan.</div><div class="grow"></div><button class="btn primary" data-a="reopen" data-id="'+s.id+'">Isi formulir</button></div>'
     :'<div class="banner info"><div><strong>Laporan sedang kami periksa.</strong>Dikirim '+dfmt(s.sentAt)+'. Estimasi hasil: '+Math.max(0,slaLeft(s))+' hari kerja lagi.</div></div>';
  }
  const eff=s.status==='qc'&&client?'review':s.status;
  const head1='<div class="page-h"><button class="btn sm" data-a="back">'+ic('back')+'Kembali</button><div><h1>'+esc(c.nama)+' · '+periodLabel(s.t)+'</h1><p>'+(s.sentAt?'Dikirim '+dfmt(s.sentAt):'Belum dikirim')+(s.rev>1?' · revisi ke-'+s.rev:'')+(client?'':' · reviewer '+esc(UNAME(reviewerOf(s)))+(mitraOf(s)?' · mitra '+esc(UNAME(mitraOf(s))):''))+'</p></div><div class="grow"></div>'+subPill(s,client)+'</div>';
  let actionsHtml='';
  if(rv&&s.status!=='disetujui'){
    const open=s.findings.filter(f=>f.st!=='ditutup').length, term=s.findings.filter(f=>f.st==='terbuka').length;
    actionsHtml='<div class="card"><div class="card-h"><h2>Keputusan review</h2></div><div class="card-b" style="display:flex;flex-direction:column;gap:12px">'+
     (s.status==='menunggu'?'<div class="row"><button class="btn primary" data-a="startrev" data-id="'+s.id+'">Mulai review</button></div>':'')+
     (s.status==='review'?'<p class="note" style="margin:0">Catat temuan di panel kanan, lalu pilih hasil. Status sesuai hanya bisa dipilih bila semua temuan sudah ditutup.</p><div class="row"><button class="btn primary" data-a="verdict" data-id="'+s.id+'" data-v="sesuai"'+(open?' title="Masih ada temuan yang belum ditutup"':'')+'>Selesaikan: sesuai</button><button class="btn danger" data-a="verdict" data-id="'+s.id+'" data-v="perbaikan">Selesaikan: minta perbaikan'+(term?' ('+term+' temuan)':'')+'</button></div>':'')+
     (s.status==='perbaikan'?'<p class="note" style="margin:0">Menunggu tanggapan klien. Anda dapat menutup temuan yang sudah dijawab.</p>':'')+
     (s.status==='qc'?(canRelease()?'<div class="banner warn"><div><strong>Hasil mitra menunggu rilis mutu.</strong>Usulan: '+(s.pend==='sesuai'?'sesuai':'minta perbaikan')+'. Periksa temuan lalu rilis ke klien.</div></div><div class="row"><button class="btn primary" data-a="release" data-id="'+s.id+'">Rilis hasil ke klien</button><button class="btn" data-a="retmitra" data-id="'+s.id+'">Kembalikan ke mitra</button></div>':'<p class="note" style="margin:0">Hasil Anda sudah dikirim dan menunggu rilis oleh tim Review & Mutu.</p>'):'')+
     (!mitra&&(isSU()||(isTim()&&me().div==='review'))?'<label class="f">Tugaskan ke mitra<select data-change="submitra" data-id="'+s.id+'"><option value="">Dikerjakan tim internal</option>'+USERS.filter(x=>x.role==='mitra').map(x=>'<option value="'+x.id+'"'+(mitraOf(s)===x.id?' selected':'')+'>'+esc(x.nama)+'</option>').join('')+'</select></label>':'')+
     '</div></div>';
  }
  const checksCard=s.fig&&!client?'<div class="card"><div class="card-h"><h2>Pra-periksa otomatis</h2><span class="grow"></span>'+(checkSummary(checks)==='ok'?pill('good','Lolos'):checkSummary(checks)==='warn'?pill('warn','Ada peringatan'):pill('bad','Ada ketidaksesuaian'))+'</div>'+checks.map(x=>'<div class="ck '+x.lvl+'"><i>'+(x.lvl==='ok'?'✓':x.lvl==='warn'?'!':'×')+'</i><div><strong>'+x.t+'</strong><small>'+esc(x.d)+'</small></div></div>').join('')+'</div>':'';
  const fcard='<div class="card"><div class="card-h"><h2>'+(client?'Laporan yang Anda kirim':'Angka yang dilaporkan klien')+'</h2></div>'+figTable(s)+(s.catatan?'<div class="card-b" style="border-top:1px solid var(--line)"><div class="note">Catatan klien</div>'+esc(s.catatan)+'</div>':'')+'</div>';
  const files='<div class="card"><div class="card-h"><h2>Berkas</h2></div><div class="list">'+DOCREQ.map(d=>'<div class="item"><div class="grow"><strong>'+d.nama+'</strong><small>'+(s.files[d.k]?esc(s.files[d.k]):'Belum diunggah')+'</small></div>'+(s.files[d.k]?pill('good','Ada'):pill('bad','Kurang'))+'</div>').join('')+'</div></div>';
  const fndHtml='<div class="card"><div class="card-h"><h2>Temuan review</h2><span class="grow"></span>'+pill('mute',String(finds.length))+'</div>'+
    (finds.length?finds.map(f=>{
      let acts='';
      if(client&&f.st==='terbuka'&&s.status==='perbaikan') acts='<label class="f">Tanggapan Anda<textarea data-input="reply" data-fid="'+f.id+'" placeholder="Jelaskan perbaikan yang dilakukan atau berikan klarifikasi.">'+esc(S.reply[s.id+f.id]||'')+'</textarea></label><div class="row"><button class="btn primary sm" data-a="answer" data-id="'+s.id+'" data-fid="'+f.id+'">Kirim tanggapan</button></div>';
      if(rv&&f.st==='dijawab') acts='<div class="row"><button class="btn sm primary" data-a="closef" data-id="'+s.id+'" data-fid="'+f.id+'">Tutup temuan</button><button class="btn sm" data-a="reopenf" data-id="'+s.id+'" data-fid="'+f.id+'">Minta tanggapan ulang</button></div>';
      if(rv&&f.st==='terbuka'&&s.status==='review') acts='<div class="row"><button class="btn sm" data-a="closef" data-id="'+s.id+'" data-fid="'+f.id+'">Tutup temuan</button></div>';
      return '<div class="fnd"><div class="meta">'+sevPill(f.sev)+pill('mute',esc(f.area))+fndPill(f)+'</div><div>'+esc(f.teks)+'</div>'+(f.resp?'<div class="resp"><div class="note">Tanggapan klien</div>'+esc(f.resp)+'</div>':'')+acts+'</div>'; }).join(''):'<div class="empty">'+(client&&!showF?'Temuan ditampilkan setelah hasil review dirilis.':'Belum ada temuan.')+'</div>')+
    (rv&&s.status==='review'?'<div class="card-b" style="border-top:1px solid var(--line);display:flex;flex-direction:column;gap:10px"><div class="row"><label class="f">Area<select data-change="fnd_area">'+AREAS.map(a=>'<option'+(S.fnd.area===a?' selected':'')+'>'+a+'</option>').join('')+'</select></label><label class="f">Tingkat<select data-change="fnd_sev"><option value="tinggi"'+(S.fnd.sev==='tinggi'?' selected':'')+'>Tinggi</option><option value="sedang"'+(S.fnd.sev==='sedang'?' selected':'')+'>Sedang</option><option value="rendah"'+(S.fnd.sev==='rendah'?' selected':'')+'>Rendah</option></select></label></div><label class="f">Uraian temuan dan saran perbaikan<textarea data-input="fnd_teks" placeholder="Tulis temuan yang jelas dan dapat ditindaklanjuti klien.">'+esc(S.fnd.teks)+'</textarea></label><div class="row"><button class="btn" data-a="addfind" data-id="'+s.id+'">'+ic('plus')+'Tambah temuan</button></div></div>':'')+'</div>';
  const hist=s.hist.filter(h=>!client||!h.internal);
  const tl='<div class="card"><div class="card-h"><h2>Riwayat'+(client?'':' dan catatan internal')+'</h2></div><div class="tl">'+(hist.length?hist.map(h=>'<div class="e"><time>'+h.ts+'</time><div class="'+(h.internal?'int':'')+'"><strong>'+esc(h.who)+'</strong>'+(h.internal?' '+pill('warn','Internal'):'')+'<div>'+esc(h.text)+'</div></div></div>').join(''):'<div class="empty">Belum ada riwayat.</div>')+'</div>'+
    (!client?'<div class="card-b" style="border-top:1px solid var(--line)"><label class="f">Catatan internal (tidak terlihat klien)<textarea data-input="note" placeholder="Catatan untuk tim atau mitra.">'+esc(S.note)+'</textarea></label><div class="row" style="margin-top:8px"><button class="btn sm" data-a="addnote" data-id="'+s.id+'">Simpan catatan</button></div></div>':'')+'</div>';
  return head1+banner+'<div class="card"><div class="card-b">'+stepsFor({status:eff})+'</div></div>'+
    '<div class="grid g2"><div style="display:flex;flex-direction:column;gap:16px">'+(client?fndHtml+fcard:actionsHtml+fcard+checksCard+files)+'</div><div style="display:flex;flex-direction:column;gap:16px">'+(client?tl:fndHtml+tl)+'</div></div>';
}

/* ---------- antrean review ---------- */
function vAntrean(){
  const filters=[['semua','Semua'],['menunggu','Menunggu'],['review','Direview'],['qc','Rilis mutu'],['perbaikan','Perbaikan klien'],['disetujui','Sesuai'],['belum','Belum kirim']];
  let rows=mySubs().filter(s=>S.qF==='semua'||s.status===S.qF).sort((a,b)=>(b.t-a.t)||a.cid.localeCompare(b.cid));
  const mitra=isMitra();
  return head(mitra?'Daftar review':'Antrean review','Seluruh laporan klien menurut status. Klik baris untuk membuka ruang review.')+
  '<div class="card"><div class="card-h"><div class="filters">'+filters.map(([k,l])=>'<button data-a="qf" data-k="'+k+'" aria-pressed="'+(S.qF===k)+'">'+l+'</button>').join('')+'</div></div><div class="tbl-wrap"><table><thead><tr><th>Klien</th><th>Periode</th><th>Status</th><th>Dikirim</th><th>SLA</th><th>Pra-periksa</th><th>Reviewer</th></tr></thead><tbody>'+
  (rows.length?rows.map(s=>{ const ck=s.fig?checkSummary(runChecks(s.fig,s.files,prevFig(s.cid,s.t),s.catatan)):null; const sl=s.sentAt&&['menunggu','review'].includes(s.status)?slaLeft(s):null;
    return '<tr class="click" data-a="sub" data-id="'+s.id+'" tabindex="0"><td><strong>'+esc(cl(s.cid).nama)+'</strong></td><td>'+periodLabel(s.t)+'</td><td>'+subPill(s,false)+'</td><td class="num">'+(s.sentAt?dfmt(s.sentAt):'–')+'</td><td>'+(sl===null?'–':sl<0?pill('bad','Lewat '+(-sl)+' hari'):pill('mute','Sisa '+sl+' hari'))+'</td><td>'+(ck==='ok'?pill('good','Lolos'):ck==='warn'?pill('warn','Peringatan'):ck==='err'?pill('bad','Gagal'):'–')+'</td><td>'+esc(UNAME(mitraOf(s)||reviewerOf(s)))+(mitraOf(s)?' <span class="note">(mitra)</span>':'')+'</td></tr>'; }).join(''):'<tr><td colspan="7" class="empty">Tidak ada data.</td></tr>')+'</tbody></table></div></div>';
}

/* ---------- klien (staf) ---------- */
function vKlien(){
  const list=myClients();
  return head(isSU()?'Klien':'Klien saya','Daftar klien beserta status periode terkini, penugasan, dan tagihan.',isSU()?'<button class="btn primary" data-a="newclient">'+ic('plus')+'Tambah klien</button>':'')+
  '<div class="card"><div class="tbl-wrap"><table><thead><tr><th>Klien</th><th>Paket</th><th>Layanan yang digunakan</th><th>PIC review</th><th>Periode Okt 2026</th><th class="r">Tagihan terbuka</th></tr></thead><tbody>'+
  list.map(c=>{ const s=SUBS.find(x=>x.cid===c.id&&x.t===10), e=engOf(c.id,'S01'), inv=INVOICES.filter(i=>i.cid===c.id&&(i.status==='terkirim'||i.status==='verifikasi')).reduce((n,i)=>n+invTot(i).total,0);
   return '<tr class="click" data-a="cd" data-id="'+c.id+'" tabindex="0"><td><strong>'+esc(c.nama)+'</strong><div class="note">'+esc(c.usaha)+' · sejak '+c.sejak+'</div></td><td>'+pill('info',esc(pk(c.paket).nama))+(c.paketNext?'<div class="note">→ '+esc(pk(c.paketNext.id).nama)+' '+dfmt(c.paketNext.berlaku)+'</div>':'')+'</td><td>'+chipsOf(c)+'</td><td>'+(e?esc(UNAME(e.pic)):'–')+'</td><td>'+(s?subPill(s,false):'–')+'</td><td class="r num">'+(isAdminDiv()||isSU()?(inv?rp(inv):'–'):'')+'</td></tr>'; }).join('')+'</tbody></table></div></div>';
}
function vKlienDetailOld(c){
  const eng=ENG.filter(e=>e.cid===c.id), subs=SUBS.filter(s=>s.cid===c.id).sort((a,b)=>b.t-a.t), su=isSU();
  const sel=(cur,opts,attrs)=>'<select '+attrs+'>'+opts.map(o=>'<option value="'+o[0]+'"'+(cur===o[0]?' selected':'')+'>'+esc(o[1])+'</option>').join('')+'</select>';
  return '<div class="page-h"><button class="btn sm" data-a="cdback">'+ic('back')+'Semua klien</button><div><h1>'+esc(c.nama)+'</h1><p>'+esc(c.usaha)+' · NPWP '+esc(c.npwp)+' · '+esc(pk(c.paket).nama)+'</p></div><div class="grow"></div>'+
   (isAdminDiv()?'<button class="btn" data-a="invnew" data-cid="'+c.id+'">'+ic('receipt')+'Buat invoice</button>':'')+'<button class="btn" data-a="remsend" data-cid="'+c.id+'">'+ic('bell')+'Kirim pengingat</button><button class="btn" data-a="openmsg" data-cid="'+c.id+'">'+ic('chat')+'Buka pesan</button>'+(canFin(c.id)?'<button class="btn" data-a="akopen" data-cid="'+c.id+'">'+ic('trend')+'Akuntansi</button>':'')+'</div>'+
  '<div class="grid g2e"><div class="card"><div class="card-h"><h2>Layanan aktif dan penugasan</h2></div><div class="tbl-wrap"><table><thead><tr><th>Layanan</th><th>Divisi</th><th>PIC tim</th><th>Mitra</th></tr></thead><tbody>'+eng.map(e=>'<tr><td>'+esc(svc(e.sid).nama)+'</td><td>'+DIV[svc(e.sid).div]+'</td><td>'+(su?sel(e.pic,USERS.filter(x=>x.role==='tim').map(x=>[x.id,x.nama]),'data-change="engpic" data-id="'+e.id+'"'):esc(UNAME(e.pic)))+'</td><td>'+(su?sel(e.mitra||'',[['','Tidak ada']].concat(USERS.filter(x=>x.role==='mitra').map(x=>[x.id,x.nama])),'data-change="engmitra" data-id="'+e.id+'"'):esc(e.mitra?UNAME(e.mitra):'-'))+'</td></tr>').join('')+'</tbody></table></div></div>'+
  '<div class="card"><div class="card-h"><h2>Pengaturan klien</h2></div><div class="card-b" style="display:flex;flex-direction:column;gap:12px"><label style="display:flex;gap:8px;align-items:center"><input type="checkbox" data-change="caltoggle" data-cid="'+c.id+'"'+(c.kalender?' checked':'')+(su||isTim()?'':' disabled')+'> Kalender kepatuhan dan pengingat otomatis aktif</label><dl class="kv"><dt>Kontak klien</dt><dd>'+esc(USERS.filter(x=>x.cid===c.id).map(x=>x.nama+' ('+x.jabatan+')').join(', ')||'-')+'</dd><dt>Klien sejak</dt><dd>'+c.sejak+'</dd></dl></div></div></div>'+
  (su?'<div class="card"><div class="card-h"><h2>Akses data keuangan</h2><span class="grow"></span><span class="note">Siapa yang boleh melihat laporan keuangan klien ini</span></div><div class="card-b" style="display:flex;gap:18px;flex-wrap:wrap"><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" checked disabled> Manda (MD)</label>'+USERS.filter(x=>x.role==='tim'||x.role==='mitra').map(x=>'<label style="display:flex;gap:6px;align-items:center"><input type="checkbox" data-change="finacc" data-cid="'+c.id+'" data-uid="'+x.id+'"'+(c.finAccess.includes(x.id)?' checked':'')+'> '+esc(x.nama)+'</label>').join('')+'</div></div>':'')+'<div class="card"><div class="card-h"><h2>Riwayat periode</h2></div><div class="list">'+subs.map(s=>'<button class="item itemlink" data-a="sub" data-id="'+s.id+'"><div class="grow"><strong>'+periodLabel(s.t)+'</strong><small>'+(s.sentAt?'Dikirim '+dfmt(s.sentAt):'Belum dikirim')+' · reviewer '+esc(UNAME(mitraOf(s)||reviewerOf(s)))+'</small></div>'+subPill(s,false)+'</button>').join('')+'</div></div>';
}

function vKlienDetail(c){
  const html=vKlienDetailOld(c), k=html.indexOf('<div class="grid g2e">'), headH=html.slice(0,k), body=html.slice(k);
  const tabs='<div class="tabs" role="tablist" style="align-self:flex-start">'+[['ringkas','Ringkasan'],['paket','Paket dan layanan'],['keuangan','Keuangan'],['tagihan','Tagihan MandaGroup']].map(([t,l])=>'<button role="tab" aria-selected="'+(S.cdt===t)+'" data-a="cdt" data-k="'+t+'">'+l+'</button>').join('')+'</div>';
  const content=S.cdt==='paket'?vKlienPaket(c):S.cdt==='keuangan'?vFinKlien(c):S.cdt==='tagihan'?vKlienTagihan(c):body;
  return headH+tabs+content;
}
/* ---------- layanan ---------- */
function vLayanan(){
  const tab=(k,l)=>'<button role="tab" aria-selected="'+(S.ltab===k)+'" data-a="ltab" data-k="'+k+'">'+l+'</button>';
  let body;
  if(S.ltab==='katalog') body='<div class="tbl-wrap"><table><thead><tr><th>Kode</th><th>Layanan</th><th>Divisi</th><th>Siklus</th><th class="r">Harga dasar</th><th>SLA</th><th>Klien aktif</th><th>Tim yang menangani</th></tr></thead><tbody>'+SERVICES.map(s=>{ const es=ENG.filter(e=>e.sid===s.id); return '<tr><td class="mono">'+s.id+'</td><td><strong>'+esc(s.nama)+'</strong><div class="note" style="max-width:340px">'+esc(s.desc)+'</div></td><td>'+DIV[s.div]+'</td><td>'+s.siklus+'</td><td class="r num">'+rp(s.harga)+'</td><td>'+s.sla+' hari kerja</td><td class="num">'+es.length+'</td><td>'+esc([...new Set(es.map(e=>UNAME(e.pic)))].join(', ')||'-')+'</td></tr>'; }).join('')+'</tbody></table></div>';
  else if(S.ltab==='paket') body=vPaketCat();
  else body='<div class="tbl-wrap"><table><thead><tr><th>Klien</th><th>Layanan</th><th>PIC tim</th><th>Mitra</th></tr></thead><tbody>'+ENG.map(e=>'<tr><td>'+esc(cl(e.cid).nama)+'</td><td>'+esc(svc(e.sid).nama)+'</td><td><select data-change="engpic" data-id="'+e.id+'">'+USERS.filter(x=>x.role==='tim').map(x=>'<option value="'+x.id+'"'+(e.pic===x.id?' selected':'')+'>'+esc(x.nama)+' · '+DIV[x.div]+'</option>').join('')+'</select></td><td><select data-change="engmitra" data-id="'+e.id+'"><option value="">Tidak ada</option>'+USERS.filter(x=>x.role==='mitra').map(x=>'<option value="'+x.id+'"'+(e.mitra===x.id?' selected':'')+'>'+esc(x.nama)+'</option>').join('')+'</select></td></tr>').join('')+'</tbody></table></div>';
  return head('Layanan dan penugasan','Katalog layanan dan penugasan tim serta mitra untuk setiap klien. Perubahan langsung berlaku pada antrean review.')+'<div class="card"><div class="card-h"><div class="tabs" role="tablist">'+tab('katalog','Katalog layanan')+tab('paket','Paket bulanan')+tab('penugasan','Penugasan')+'</div></div>'+body+'</div>';
}

/* ===================== AKUNTANSI (v7): model dan perhitungan ===================== */
const T0=7, AKM=[7,8,9,10,11], BASEM=10;
/* ---------- keuangan Akuntara (internal perusahaan, hanya Managing Director) ---------- */
const KXCAT={gaji:['Gaji dan honor tim','op'],mitra:['Honor mitra pekerjaan','op'],sewa:['Sewa kantor dan coworking','op'],sw:['Perangkat lunak dan hosting','op'],net:['Internet dan utilitas','op'],mkt:['Pemasaran dan iklan','op'],adm:['Legalitas dan administrasi','op'],lain:['Beban lain-lain','op'],pph:['PPh Final 0,5%','pph'],ppn:['Setoran PPN','ppn'],alat:['Pembelian peralatan','alat']};
const KX={modal:CONFIG.modalAwal};
const KXEXP=[]; let KXSEQ=0;
function kxAdd(tgl,kat,ket,jml,user){ KXEXP.push({id:'x'+(++KXSEQ),tgl,kat,ket,jml:Math.round(jml),user:!!user}); }
function kxRev(m){ let r=0; INVOICES.forEach(i=>{ if(i.status==='draft'||i.tgl.getMonth()+1!==m) return; const t=invTot(i); if(t.jenis==='penuh') r+=t.dpp; else if(t.jenis==='pelunasan') r+=t.kontrak; }); return r; }
function kxPpn(m){ let r=0; INVOICES.forEach(i=>{ if(i.status!=='draft'&&i.tgl.getMonth()+1===m) r+=invTot(i).ppn; }); return r; }

function kxCalc(){
  const R={}, byC={}; CLIENTS.forEach(c=>byC[c.id]=0);
  let kas=KX.modal, piut=0, alat=0, ppnL=0, um=0, ret=0;
  AKM.forEach(m=>{
    const r={m,revB:0,revP:0,umD:0,ppnOut:0,issued:0,paid:0,op:{},opT:0,pph:0,ppnPaid:0,alat:0};
    INVOICES.forEach(i=>{ if(i.status==='draft') return; const t=invTot(i);
      if(i.tgl.getMonth()+1===m){ r.issued+=t.total; r.ppnOut+=t.ppn;
        if(t.jenis==='penuh'){ r.revB+=t.dpp; byC[i.cid]+=t.dpp; } else if(t.jenis==='dp'){ r.umD+=t.dpp; } else { r.revP+=t.kontrak; r.umD-=t.dpBase; byC[i.cid]+=t.kontrak; } }
      if(i.status==='lunas'&&i.tglBayar&&i.tglBayar.getMonth()+1===m) r.paid+=t.total; });
    KXEXP.forEach(x=>{ if(x.tgl.getMonth()+1!==m) return; const k=KXCAT[x.kat][1]; if(k==='op'){ r.op[x.kat]=(r.op[x.kat]||0)+x.jml; r.opT+=x.jml; } else if(k==='pph') r.pph+=x.jml; else if(k==='ppn') r.ppnPaid+=x.jml; else r.alat+=x.jml; });
    r.rev=r.revB+r.revP; r.labaU=r.rev-r.opT; r.laba=r.labaU-r.pph; r.kas0=kas; r.kasIn=r.paid; r.kasOut=r.opT+r.pph+r.ppnPaid+r.alat;
    kas+=r.kasIn-r.kasOut; piut+=r.issued-r.paid; alat+=r.alat; ppnL+=r.ppnOut-r.ppnPaid; um+=r.umD; ret+=r.laba; r.kas1=kas;
    r.bs={kas,piut,alat,aset:kas+piut+alat,ppn:ppnL,um,liab:ppnL+um,modal:KX.modal,ret,eq:KX.modal+ret}; r.bs.diff=r.bs.aset-(r.bs.liab+r.bs.eq);
    R[m]=r; });
  return {R,byC};
}
function kxAging(){ const b=[['Belum jatuh tempo',0],['Terlambat 1 sampai 30 hari',0],['Terlambat 31 sampai 60 hari',0],['Terlambat lebih dari 60 hari',0]]; INVOICES.filter(i=>i.status==='terkirim'||i.status==='verifikasi').forEach(i=>{ const late=-daysTo(i.due); b[late<=0?0:late<=30?1:late<=60?2:3][1]+=invTot(i).total; }); return b; }
const nz=n=>n<0?'('+num(n)+')':num(n);
const kxHeads=()=>AKM.map(m=>BLN3[m-1]+(m===11?' *':''));
function kxTable(rows,tot){ const hs=kxHeads(), cs=hs.length+(tot?2:1);
  return '<div class="tbl-wrap"><table class="kxt"><thead><tr><th></th>'+hs.map(h=>'<th class="r">'+h+'</th>').join('')+(tot?'<th class="r">Total</th>':'')+'</tr></thead><tbody>'+rows.map(r=>{ if(r.sec) return '<tr class="sec"><td colspan="'+cs+'">'+r.sec+'</td></tr>'; const t=r.t===undefined?r.v.reduce((a,x)=>a+x,0):r.t;
    return '<tr'+(r.b?' class="tot"':'')+'><td>'+(r.ind?'<span style="padding-left:14px">'+r.l+'</span>':r.l)+'</td>'+r.v.map(x=>'<td class="r num">'+nz(x)+'</td>').join('')+(tot?'<td class="r num">'+nz(t)+'</td>':'')+'</tr>'; }).join('')+'</tbody></table></div><p class="note" style="margin:10px 4px 0">* Bulan berjalan, belum tutup buku. Seluruh angka adalah data contoh untuk demo.</p>'; }
function kxLR(K){ const A=m=>AKM.map(x=>K.R[x][m]); const rows=[{sec:'Pendapatan'},{l:'Jasa langganan dan layanan bulanan',ind:1,v:A('revB')},{l:'Jasa berbasis proyek (DP dan pelunasan)',ind:1,v:A('revP')},{l:'Total pendapatan',b:1,v:A('rev')},{sec:'Beban operasional'}];
  Object.keys(KXCAT).filter(k=>KXCAT[k][1]==='op').forEach(k=>{ const v=AKM.map(m=>K.R[m].op[k]||0); if(v.some(x=>x)) rows.push({l:KXCAT[k][0],ind:1,v}); });
  rows.push({l:'Total beban operasional',b:1,v:A('opT')},{l:'Laba usaha',b:1,v:A('labaU')},{l:'Beban PPh Final 0,5%',ind:1,v:A('pph')},{l:'Laba bersih',b:1,v:A('laba')});
  const mg=AKM.map(m=>K.R[m].rev?Math.round(K.R[m].laba/K.R[m].rev*1000)/10:0); rows.push({l:'Margin laba bersih (%)',v:mg,t:(()=>{ const rv=AKM.reduce((a,m)=>a+K.R[m].rev,0), lb=AKM.reduce((a,m)=>a+K.R[m].laba,0); return rv?Math.round(lb/rv*1000)/10:0; })()});
  return rows; }
function kxKas(K){ const A=m=>AKM.map(x=>K.R[x][m]), N=m=>AKM.map(x=>-K.R[x][m]); const net=AKM.map(m=>K.R[m].kasIn-K.R[m].opT-K.R[m].pph-K.R[m].ppnPaid);
  return [{sec:'Aktivitas operasi'},{l:'Penerimaan dari klien',ind:1,v:A('kasIn')},{l:'Pembayaran beban operasional',ind:1,v:N('opT')},{l:'Pembayaran PPh Final',ind:1,v:N('pph')},{l:'Setoran PPN',ind:1,v:N('ppnPaid')},{l:'Kas bersih dari operasi',b:1,v:net},{sec:'Aktivitas investasi'},{l:'Pembelian peralatan',ind:1,v:N('alat')},{l:'Kenaikan (penurunan) kas',b:1,v:AKM.map(m=>K.R[m].kasIn-K.R[m].kasOut)},{l:'Saldo kas awal',v:A('kas0'),t:K.R[AKM[0]].kas0},{l:'Saldo kas akhir',b:1,v:A('kas1'),t:K.R[AKM[AKM.length-1]].kas1}]; }
function kxNeraca(K){ const B=k=>AKM.map(m=>K.R[m].bs[k]); return [{sec:'Aset'},{l:'Kas dan bank',ind:1,v:B('kas')},{l:'Piutang usaha',ind:1,v:B('piut')},{l:'Peralatan (tanpa penyusutan)',ind:1,v:B('alat')},{l:'Total aset',b:1,v:B('aset')},{sec:'Liabilitas'},{l:'Utang PPN',ind:1,v:B('ppn')},{l:'Uang muka klien (DP)',ind:1,v:B('um')},{l:'Total liabilitas',b:1,v:B('liab')},{sec:'Ekuitas'},{l:'Modal disetor',ind:1,v:B('modal')},{l:'Laba ditahan dan berjalan',ind:1,v:B('ret')},{l:'Total ekuitas',b:1,v:B('eq')},{l:'Total liabilitas dan ekuitas',b:1,v:AKM.map(m=>K.R[m].bs.liab+K.R[m].bs.eq)}]; }
function kxCsv(){ const K=kxCalc(), L=[]; const add=(...a)=>L.push(a.map(x=>'"'+String(x).replace(/"/g,'""')+'"').join(','));
  const dump=(title,rows,tot)=>{ add(title); add('',...kxHeads().map(h=>h.replace(' *',' (berjalan)')),...(tot?['Total']:[])); rows.forEach(r=>{ if(r.sec){ add(r.sec); return; } const t=r.t===undefined?r.v.reduce((a,x)=>a+x,0):r.t; add(r.l,...r.v,...(tot?[t]:[])); }); L.push(''); };
  add(CONFIG.entity); add('Laporan keuangan internal MandaGroup 2026 (data contoh)'); L.push('');
  dump('LABA RUGI (Rp)',kxLR(K),true); dump('ARUS KAS (Rp)',kxKas(K),true); dump('NERACA (Rp)',kxNeraca(K),false); return L.join('\r\n'); }
function vKakt(){
  if(!isSU()) return head('Keuangan MandaGroup','Laporan keuangan internal perusahaan.')+locked('Hanya untuk Managing Director','Laporan keuangan perusahaan bersifat rahasia dan hanya dapat dibuka oleh Managing Director.');
  const K=kxCalc(), tab=S.kt, TABS=[['ringkas','Ringkasan'],['masuk','Pemasukan'],['keluar','Pengeluaran'],['lr','Laba rugi'],['kas','Arus kas'],['neraca','Neraca'],['pajak','Pajak']];
  const tabs='<div class="tabs" role="tablist">'+TABS.map(([k,l])=>'<button role="tab" aria-selected="'+(tab===k)+'" data-a="kt" data-k="'+k+'">'+l+'</button>').join('')+'</div>';
  const exp='<div class="row" style="flex-wrap:nowrap;gap:8px"><button class="btn" data-a="kxcsv">Unduh CSV</button><button class="btn" data-a="kxprint">Cetak atau simpan PDF</button></div>';
  let body='', act='';
  const kop=t=>'<div class="card card-pad kop" style="margin-bottom:12px"><span class="eyebrow">'+esc(CONFIG.entity)+'</span><h2 style="margin:2px 0">'+t+'</h2><p class="note" style="margin:0">Juli sampai November 2026 · dalam Rupiah · data contoh untuk demo</p></div>';
  if(tab==='ringkas'){ const m=S.ktp, r=K.R[m], tl=AKM.reduce((a,x)=>a+K.R[x].rev,0);
    const mrr=CLIENTS.reduce((a,c)=>a+(pk(c.paket)?pk(c.paket).harga:0),0), ag=kxAging(), agT=ag.reduce((a,x)=>a+x[1],0);
    const top=CLIENTS.map(c=>({c,v:K.byC[c.id]})).sort((a,b)=>b.v-a.v);
    act='<label class="f" style="min-width:170px">Periode<select data-change="ktp">'+AKM.map(x=>'<option value="'+x+'"'+(x===m?' selected':'')+'>'+BULAN[x-1]+' 2026'+(x===11?' (berjalan)':'')+'</option>').join('')+'</select></label>';
    body=finStrip([['white','trend','Pendapatan',short(r.rev),periodLabel(m)+(m===11?' · belum tutup':'')],[r.laba>=0?'mint':'peach','wallet','Laba bersih',short(r.laba),r.rev?(r.laba/r.rev*100).toFixed(1).replace('.',',')+'% dari pendapatan':'Belum ada pendapatan'],['blue','wallet','Saldo kas akhir',short(r.kas1),'Awal '+short(r.kas0)],[agT?'white':'mint','receipt','Piutang usaha',short(r.bs.piut),'Posisi akhir '+BLN3[m-1]]])+
    '<div class="grid g2e"><section class="card card-pad"><span class="eyebrow">KINERJA</span><h2 style="margin:2px 0 10px">Pendapatan dan beban per bulan</h2>'+bars(AKM.map(x=>BLN3[x-1]),[{name:'Pendapatan',color:'var(--brand)',vals:AKM.map(x=>K.R[x].rev)},{name:'Beban dan pajak',color:'var(--copper)',vals:AKM.map(x=>K.R[x].opT+K.R[x].pph)}])+'<p class="note" style="margin:8px 0 0">Bulan November masih berjalan: beban gaji dan honor belum dicatat.</p></section>'+
    '<section class="card card-pad"><span class="eyebrow">PIUTANG</span><h2 style="margin:2px 0 10px">Umur piutang saat ini</h2><dl class="kv" style="grid-template-columns:1fr auto">'+ag.map(x=>'<dt>'+x[0]+'</dt><dd class="num">'+rp(x[1])+'</dd>').join('')+'<dt><strong>Total piutang</strong></dt><dd class="num"><strong>'+rp(agT)+'</strong></dd></dl></section></div>'+
    '<div class="grid g2e"><section class="card"><div class="card-h"><h2>Pendapatan per klien</h2><span class="grow"></span><span class="note">Jul sampai Nov 2026</span></div><div class="tbl-wrap"><table><thead><tr><th>Klien</th><th class="r">Pendapatan</th><th class="r">Porsi</th></tr></thead><tbody>'+top.map(x=>'<tr><td>'+esc(x.c.nama)+'</td><td class="r num">'+num(x.v)+'</td><td class="r num">'+(tl?(x.v/tl*100).toFixed(1).replace('.',','):'0')+'%</td></tr>').join('')+'</tbody></table></div></section>'+
    '<section class="card card-pad"><span class="eyebrow">PENDAPATAN BERULANG</span><h2 style="margin:2px 0 10px">Paket langganan aktif</h2><dl class="kv" style="grid-template-columns:1fr auto"><dt>Pendapatan berulang bulanan (MRR)</dt><dd class="num">'+rp(mrr)+'</dd><dt>Proyeksi tahunan (ARR)</dt><dd class="num">'+rp(mrr*12)+'</dd><dt>Klien berlangganan</dt><dd class="num">'+CLIENTS.length+'</dd><dt>Rata-rata per klien</dt><dd class="num">'+rp(Math.round(mrr/Math.max(1,CLIENTS.length)))+'</dd></dl><p class="note" style="margin:8px 0 0">Dihitung dari harga paket klien aktif, belum termasuk layanan tambahan dan proyek.</p></section></div>'; }
  else if(tab==='masuk'){ const rows=INVOICES.filter(i=>i.status!=='draft').sort((a,b)=>b.tgl-a.tgl||b.no.localeCompare(a.no)), all=rows.reduce((a,i)=>a+invTot(i).total,0), got=rows.filter(i=>i.status==='lunas').reduce((a,i)=>a+invTot(i).total,0), late=rows.filter(i=>i.status==='terkirim'&&daysTo(i.due)<0).reduce((a,i)=>a+invTot(i).total,0);
    act='<button class="btn primary" data-a="invnew">'+ic('plus')+'Buat invoice</button>';
    body=finStrip([['white','receipt','Total ditagihkan',short(all),rows.length+' invoice terbit'],['mint','filecheck','Sudah diterima',short(got),'Kas masuk dari klien'],['white','hourglass','Belum dibayar',short(all-got),'Termasuk menunggu verifikasi'],[late?'peach':'white','alert','Terlambat',short(late),late?'Perlu pengingat tagihan':'Tidak ada']])+
    '<div class="card"><div class="tbl-wrap"><table><thead><tr><th>No. invoice</th><th>Klien</th><th>Tanggal</th><th>Jatuh tempo</th><th>Dibayar</th><th class="r">Total</th><th>Status</th></tr></thead><tbody>'+rows.map(i=>'<tr class="click" data-a="invopen" data-no="'+i.no+'" tabindex="0"><td class="mono">'+i.no+jenisTag(i)+'</td><td>'+esc(cl(i.cid).nama)+'</td><td class="num">'+dfmt(i.tgl)+'</td><td class="num">'+dfmt(i.due)+'</td><td class="num">'+(i.tglBayar?dfmt(i.tglBayar):'–')+'</td><td class="r num"><strong>'+num(invTot(i).total)+'</strong></td><td>'+invStatus(i)+'</td></tr>').join('')+'</tbody></table></div></div>'; }
  else if(tab==='keluar'){ const rows=KXEXP.slice().sort((a,b)=>b.tgl-a.tgl), tot=KXEXP.reduce((a,x)=>a+x.jml,0);
    act='<button class="btn primary" data-a="kxnew">'+ic('plus')+'Catat pengeluaran</button>';
    const byK={}; KXEXP.forEach(x=>byK[x.kat]=(byK[x.kat]||0)+x.jml);
    body=finStrip([['white','wallet','Total pengeluaran',short(tot),'Juli sampai November 2026'],['white','team','Gaji, honor, dan mitra',short((byK.gaji||0)+(byK.mitra||0)),'Biaya tenaga kerja'],['white','receipt','Pajak dibayar',short((byK.pph||0)+(byK.ppn||0)),'PPh Final dan PPN'],['blue','gear','Peralatan',short(byK.alat||0),'Dicatat sebagai aset']])+
    '<div class="card"><div class="tbl-wrap"><table><thead><tr><th>Tanggal</th><th>Kategori</th><th>Keterangan</th><th class="r">Jumlah</th><th></th></tr></thead><tbody>'+rows.map(x=>'<tr><td class="num">'+dfmt(x.tgl)+'</td><td>'+esc(KXCAT[x.kat][0])+'</td><td>'+esc(x.ket)+'</td><td class="r num">'+num(x.jml)+'</td><td class="r">'+(x.user?'<button class="btn sm" data-a="kxdel" data-id="'+x.id+'" aria-label="Hapus">'+ic('x')+'</button>':'')+'</td></tr>').join('')+'</tbody></table></div></div>'; }
  else if(tab==='lr'||tab==='kas'||tab==='neraca'){ const T3={lr:['Laporan laba rugi',kxLR(K),true],kas:['Laporan arus kas',kxKas(K),true],neraca:['Neraca',kxNeraca(K),false]}[tab]; act=exp;
    const ok=AKM.every(m=>K.R[m].bs.diff===0);
    body='<div class="print-area">'+kop(T3[0])+'<div class="card">'+kxTable(T3[1],T3[2])+'</div>'+(tab==='neraca'?'<p style="margin:10px 4px 0">'+(ok?pill('good','Neraca seimbang pada semua periode'):pill('bad','Neraca tidak seimbang'))+'</p>':'')+'</div>'; }
  else { const A=m=>AKM.map(x=>K.R[x][m]); const ytd=AKM.reduce((a,m)=>a+K.R[m].rev,0), lim=4800000000;
    act=exp;
    body='<div class="print-area">'+kop('Ringkasan pajak')+'<div class="card">'+kxTable([{sec:'PPh Final 0,5% (dasar: pendapatan)'},{l:'Pendapatan (dasar pengenaan)',ind:1,v:A('rev')},{l:'PPh Final terutang',ind:1,v:AKM.map(m=>Math.round(K.R[m].rev*CONFIG.pph))},{l:'PPh Final dibayar pada bulan ini',ind:1,v:A('pph')},{sec:'PPN (asumsi PKP)'},{l:'PPN keluaran dari invoice',ind:1,v:A('ppnOut')},{l:'Setoran PPN',ind:1,v:A('ppnPaid')},{l:'Saldo utang PPN akhir bulan',b:1,v:AKM.map(m=>K.R[m].bs.ppn),t:K.R[AKM[AKM.length-1]].bs.ppn}],true)+'</div>'+
    '<section class="card card-pad" style="margin-top:12px"><span class="eyebrow">BATAS OMZET</span><h2 style="margin:2px 0 8px">Akumulasi pendapatan tahun berjalan</h2>'+bar(Math.min(100,ytd/lim*100),'var(--brand)')+'<p class="note" style="margin:8px 0 0">'+rp(ytd)+' dari batas peredaran bruto Rp 4.800.000.000 untuk tarif PPh Final UMKM. Angka ini estimasi untuk perencanaan.</p></section>'+
    '<p class="note" style="margin:12px 4px 0">Estimasi saja, bukan pengganti perhitungan konsultan pajak. PPh Final disetor paling lambat tanggal 15 bulan berikutnya dan PPN paling lambat akhir bulan berikutnya. Masa berlaku tarif final 0,5% bagi perseroan terbatas dibatasi 3 tahun sejak pendirian, dan PPN mengasumsikan perusahaan telah dikukuhkan sebagai PKP.</p></div>'; }
  return head('Keuangan MandaGroup','Pemasukan, pengeluaran, dan laporan keuangan '+esc(CONFIG.entity)+'. Hanya dapat dilihat Managing Director.',act)+tabs+body;
}

const COA=[
 {k:'1100',n:'Kas dan Bank',t:'aset',g:'lancar'},
 {k:'1200',n:'Piutang Usaha',t:'aset',g:'lancar'},
 {k:'1300',n:'Persediaan',t:'aset',g:'lancar'},
 {k:'1500',n:'Aset Tetap (nilai buku)',t:'aset',g:'tetap'},
 {k:'2100',n:'Utang Usaha',t:'liab',g:'lancar'},
 {k:'2200',n:'Utang Pajak',t:'liab',g:'lancar'},
 {k:'2300',n:'Utang Bank',t:'liab',g:'panjang'},
 {k:'3100',n:'Modal Disetor',t:'ekuitas'},
 {k:'3200',n:'Laba Ditahan',t:'ekuitas'},
 {k:'4100',n:'Pendapatan Usaha',t:'pend'},
 {k:'5100',n:'Harga Pokok Penjualan',t:'beban',g:'hpp'},
 {k:'6100',n:'Beban Gaji',t:'beban',g:'opx'},
 {k:'6200',n:'Beban Sewa',t:'beban',g:'opx'},
 {k:'6300',n:'Beban Operasional',t:'beban',g:'opx'}
];
const TIPE_AKUN={aset:'Aset',liab:'Liabilitas',ekuitas:'Ekuitas',pend:'Pendapatan',beban:'Beban'};
const ACCX={}, OPEN={}, TRX=[], CLOSED={}, AUDIT=[], ACCFIND=[], NOTES=[], TARGETS=[], SESI=[], BANKX=[], PROJ={}, BUDG={};
const TAXM={}, PKP={};
let TSEQ=0, AUSEQ=0, FSEQ=0, CSEQ=0, GSEQ=0, SSEQ=0, XSEQ=0;
const COAOF=cid=>COA.concat(ACCX[cid]||[]);
const accOf=(cid,k)=>COAOF(cid).find(a=>a.k===k);
const rK=n=>Math.round(n/1000)*1000;
const pc=v=>(isFinite(v)?(v*100).toFixed(1).replace('.',','):'0,0')+'%';
const sh=v=>(v<0?'-':'')+short(v);
const isClosed=(cid,t)=>!!CLOSED[cid+'|'+t];
const dayLbl=(t,d)=>String(d).padStart(2,'0')+' '+BLN3[(t-1)%12];
function addTrx(o){ TSEQ++; const x=Object.assign({id:'T'+TSEQ,by:'Sistem',ts:''},o); TRX.push(x); return x; }
function audit(cid,kat,text,who){ AUSEQ++; AUDIT.unshift({id:'A'+AUSEQ,ts:stamp(),who:who||me().nama,cid,kat,text}); }

/* ---------- perhitungan laporan dari buku besar ---------- */
function sumTrx(cid,t0,t1){ const m={}; TRX.forEach(x=>{ if(x.cid===cid&&x.t>=t0&&x.t<=t1){ (m[x.deb]=m[x.deb]||{d:0,k:0}).d+=x.jml; (m[x.kre]=m[x.kre]||{d:0,k:0}).k+=x.jml; } }); return m; }
const natOf=(a,m)=>{ const s=m[a.k]||{d:0,k:0}; return (a.t==='aset'||a.t==='beban')?s.d-s.k:s.k-s.d; };
const sumV=l=>l.reduce((s,x)=>s+x.v,0);
function buku(cid,t){
  const A=COAOF(cid), op=OPEN[cid]||{}, cum=sumTrx(cid,T0,t), mon=sumTrx(cid,t,t);
  const bal=a=>(op[a.k]||0)+natOf(a,cum);
  const bs=f=>A.filter(f).map(a=>({k:a.k,n:a.n,v:bal(a)}));
  const plr=m=>({pend:A.filter(a=>a.t==='pend').map(a=>({k:a.k,n:a.n,v:natOf(a,m)})),hpp:A.filter(a=>a.t==='beban'&&a.g==='hpp').map(a=>({k:a.k,n:a.n,v:natOf(a,m)})),opx:A.filter(a=>a.t==='beban'&&a.g!=='hpp').map(a=>({k:a.k,n:a.n,v:natOf(a,m)}))});
  const asetL=bs(a=>a.t==='aset'&&a.g!=='tetap'), asetT=bs(a=>a.t==='aset'&&a.g==='tetap');
  const liabL=bs(a=>a.t==='liab'&&a.g!=='panjang'), liabP=bs(a=>a.t==='liab'&&a.g==='panjang'), eq=bs(a=>a.t==='ekuitas');
  const lr=plr(mon), lrc=plr(cum);
  const pend=sumV(lr.pend), hpp=sumV(lr.hpp), opx=sumV(lr.opx), laba=pend-hpp-opx;
  const labac=sumV(lrc.pend)-sumV(lrc.hpp)-sumV(lrc.opx);
  const totA=sumV(asetL)+sumV(asetT), totL=sumV(liabL)+sumV(liabP), totE=sumV(eq)+labac;
  const kasA=A.find(a=>a.k==='1100'), kas1=bal(kasA), kas0=(op['1100']||0)+natOf(kasA,sumTrx(cid,T0,t-1));
  const m1=mon['1100']||{d:0,k:0};
  const accs=A.map(a=>({k:a.k,n:a.n,t:a.t,v:(a.t==='aset'||a.t==='liab'||a.t==='ekuitas')?bal(a):natOf(a,mon)}));
  return {asetL,asetT,liabL,liabP,eq,lr,lrc,pend,hpp,opx,laba,labac,totA,totL,totE,kas0,kas1,masuk:m1.d,keluar:m1.k,accs};
}
const CFG={'1500':'Investasi','2300':'Pendanaan','3100':'Pendanaan','3200':'Pendanaan'};
function cashflow(cid,t){
  const by={}; TRX.filter(x=>x.cid===cid&&x.t===t&&(x.deb==='1100'||x.kre==='1100')).forEach(x=>{
    const inn=x.deb==='1100', ctr=inn?x.kre:x.deb, grp=CFG[ctr]||'Operasional', key=grp+'|'+ctr;
    by[key]=by[key]||{grp,ctr,v:0}; by[key].v+=inn?x.jml:-x.jml; });
  return ['Operasional','Investasi','Pendanaan'].map(g=>{ const rows=Object.values(by).filter(r=>r.grp===g).map(r=>{ const a=accOf(cid,r.ctr); return {n:(r.v>=0?'Penerimaan: ':'Pembayaran: ')+(a?a.n:r.ctr),v:r.v}; }); return {g,rows,tot:sumV(rows)}; });
}
function figLedger(cid,t){ const s=buku(cid,t); return {pendapatan:s.pend,beban:s.hpp+s.opx,laba:s.laba,aset:s.totA,liab:s.totL,ekuitas:s.totE,kas0:s.kas0,masuk:s.masuk,keluar:s.keluar,kas1:s.kas1}; }

/* ---------- rasio keuangan ---------- */
const rst=(v,g,w,hi)=>hi?(v>=g?'good':v>=w?'warn':'bad'):(v<=g?'good':v<=w?'warn':'bad');
const RST={good:'Sehat',warn:'Perlu dipantau',bad:'Perlu tindakan',mute:'Belum ada data'};
function ratios(cid,t){
  const s=buku(cid,t), piu=(s.asetL.find(x=>x.k==='1200')||{v:0}).v, aL=sumV(s.asetL), lL=sumV(s.liabL), bebanKas=s.opx;
  const dso=s.pend>0?piu/s.pend*30:null, lancar=lL>0?aL/lL:null, runway=bebanKas>0?s.kas1/bebanKas:null, mg=s.pend>0?s.laba/s.pend:null;
  return [
   {k:'margin',label:'Margin laba bersih',val:mg===null?'–':pc(mg),st:mg===null?'mute':rst(mg,.10,.05,true),hint:mg===null?'Belum ada pendapatan pada periode ini':'Margin kotor '+pc((s.pend-s.hpp)/s.pend)+'. Sehat bila 10% atau lebih.',raw:mg},
   {k:'lancar',label:'Rasio lancar',val:lancar===null?'–':lancar.toFixed(2).replace('.',',')+' kali',st:lancar===null?'mute':rst(lancar,1.5,1,true),hint:'Aset lancar dibanding liabilitas lancar. Sehat bila 1,5 kali atau lebih.',raw:lancar},
   {k:'piutang',label:'Perputaran piutang',val:dso===null?'–':Math.round(dso)+' hari',st:dso===null?'mute':rst(dso,45,60,false),hint:dso===null?'Belum ada pendapatan pada periode ini':'Sekitar '+(piu>0?(s.pend*12/piu).toFixed(1).replace('.',','):'0')+' kali per tahun. Sehat bila 45 hari atau kurang.',raw:dso},
   {k:'kas',label:'Umur kas',val:runway===null?'–':runway.toFixed(1).replace('.',',')+' bulan',st:runway===null?'mute':rst(runway,3,1.5,true),hint:'Saldo kas dibanding beban operasional bulan ini (di luar HPP). Sehat bila 3 bulan atau lebih.',raw:runway}
  ];
}

/* ---------- anggaran, proyeksi ---------- */
const BUDK=['4100','5100','6100','6200','6300'];
function defBud(cid){ const s=buku(cid,9), all=[...s.lr.pend,...s.lr.hpp,...s.lr.opx], g=k=>(all.find(x=>x.k===k)||{v:0}).v; const o={}; BUDK.forEach(k=>{ o[k]=rK(g(k)*(k==='4100'?1.03:1)); }); return o; }
const budOf=cid=>BUDG[cid]||(BUDG[cid]=defBud(cid));
function budRows(cid,t){ const s=buku(cid,t), all=[...s.lr.pend,...s.lr.hpp,...s.lr.opx], b=budOf(cid);
  return BUDK.map(k=>{ const a=all.find(x=>x.k===k)||{n:k,v:0}, bud=b[k]||0, real=a.v, sel=real-bud, pct=bud?sel/bud:0, inc=k==='4100';
    const st=!bud?'mute':inc?(pct<=-.10?'bad':pct<=-.05?'warn':'good'):(pct>=.10?'bad':pct>=.05?'warn':'good'); return {k,n:a.n,bud,real,sel,pct,st,inc}; }); }
function defProj(cid){ const s=buku(cid,BASEM), fx=(s.lr.opx.filter(x=>x.k==='6100'||x.k==='6200')).reduce((a,x)=>a+x.v,0), vp=s.pend>0?(s.hpp+s.opx-fx)/s.pend*100:70; return {g:2,vp:Math.round(vp*10)/10,fx:fx}; }
const pa=cid=>PROJ[cid]||(PROJ[cid]=defProj(cid));
const SCN=[{k:'kons',n:'Konservatif',dg:-2,dvp:1,dfx:.05,color:'var(--warn)'},{k:'dasar',n:'Dasar',dg:0,dvp:0,dfx:0,color:'var(--brand)'},{k:'opt',n:'Optimis',dg:2,dvp:-1,dfx:-.03,color:'var(--good)'}];
function projRun(cid,hor,sc){ const a=pa(cid), base=buku(cid,BASEM); let rev=base.pend, kas=base.kas1; const out=[];
  for(let i=1;i<=hor;i++){ rev=rev*(1+(a.g+sc.dg)/100); const beban=rev*(a.vp+sc.dvp)/100+a.fx*(1+sc.dfx), laba=rev-beban; kas+=laba; const t=BASEM+i;
    out.push({t,lbl:BLN3[(t-1)%12]+' '+String(2026+Math.floor((t-1)/12)).slice(2),rev:Math.round(rev),beban:Math.round(beban),laba:Math.round(laba),kas:Math.round(kas)}); }
  return out; }

/* ---------- pajak (estimasi) ---------- */
function taxEst(cid,t){
  const m=buku(cid,t), c=buku(cid,t), mode=TAXM[cid]||'final', pkp=!!PKP[cid];
  const cumP=sumV(c.lrc.pend), cumL=c.labac, cumH=sumV(c.lrc.hpp);
  const pph=x=>mode==='final'?{nama:'PPh Final UMKM 0,5% dari omzet',bulan:Math.round(m.pend*.005),kum:Math.round(cumP*.005)}:{nama:'PPh Badan 22% dari laba sebelum pajak',bulan:Math.round(Math.max(0,m.laba)*.22),kum:Math.round(Math.max(0,cumL)*.22)};
  const p=pph();
  const ppn=pkp?{keluaran:Math.round(m.pend*CONFIG.ppn),masukan:Math.round(m.hpp*CONFIG.ppn),kKel:Math.round(cumP*CONFIG.ppn),kMas:Math.round(cumH*CONFIG.ppn)}:null;
  return {mode,pkp,pph:p,ppn};
}


/* sinkronkan angka laporan yang sudah dikirim dengan buku besar */
SUBS.forEach(s=>{ if(!s.fig||!RV[s.cid]||s.t<T0||s.t>10) return; const f=figLedger(s.cid,s.t);
  if(s.cid==='c1'&&s.t===10) f.ekuitas+=12500000;
  if(s.cid==='c4'&&s.t===10){ f.kas0+=3200000; f.keluar=f.kas0+f.masuk-f.kas1; }
  s.fig=f; });

/* ===================== AKUNTANSI (v7): tampilan ===================== */
Object.assign(S,{akc:null,akt:10,akb:'laporan',prhor:6,prsc:'dasar',prmet:'laba',rvk:'1200',afk:'semua',afq:'',reopen:'',afr:{},
  trxf:{t:10,d:'',jenis:'tunai',deb:'1100',kre:'4100',jml:'',ket:''},accf:{k:'',n:'',t:'beban'},afs:{k:'1200',sev:'sedang',teks:''},
  notef:{judul:'',teks:'',internal:false},tgf:{metrik:'laba',nilai:''},sesf:{tgl:'',topik:''}});
const AKF=['trxf','accf','afs','notef','tgf','sesf'];
const AKTABS=[['laporan','Laporan'],['transaksi','Transaksi'],['rasio','Rasio'],['anggaran','Anggaran'],['proyeksi','Proyeksi'],['rekon','Rekonsiliasi bank'],['review','Review akun'],['pajak','Pajak'],['catatan','Catatan konsultan'],['target','Target dan jadwal'],['periode','Tutup buku'],['audit','Jejak audit']];
const JENIS=[['tunai','Penjualan tunai atau transfer','1100','4100'],['kredit','Penjualan kredit (piutang)','1200','4100'],['terima','Penerimaan piutang dari pelanggan','1100','1200'],['hpp','Pembelian barang atau bahan tunai','5100','1100'],['hppk','Pembelian barang atau bahan kredit','5100','2100'],['bayar','Pembayaran utang usaha','2100','1100'],['gaji','Pembayaran gaji','6100','1100'],['sewa','Pembayaran sewa','6200','1100'],['ops','Beban operasional tunai','6300','1100'],['aset','Pembelian aset tetap','1500','1100'],['bank','Angsuran utang bank','2300','1100'],['modal','Setoran modal','1100','3100'],['manual','Jurnal manual','','']];
const canInputAk=cid=>canFin(cid)&&(isKlien()||isSU()||isTim());
const canStaffAk=cid=>canFin(cid)&&(isSU()||isTim());
const canReviewAk=cid=>canFin(cid)&&(isSU()||isTim()||isMitra());
const tbl=(h,r)=>'<div class="tbl-wrap"><table><thead><tr>'+h.map(x=>x[0]==='>'?'<th class="r">'+x.slice(1)+'</th>':'<th>'+x+'</th>').join('')+'</tr></thead><tbody>'+r.join('')+'</tbody></table></div>';
const emptyBox=t=>'<div class="empty">'+t+'</div>';
const lockBanner=(cid,t)=>isClosed(cid,t)?'<div class="banner info"><div><strong>'+periodLabel(t)+' sudah ditutup</strong>Transaksi dan rekonsiliasi pada periode ini terkunci. Pembukaan kembali hanya dapat dilakukan Managing Director dengan alasan tercatat.</div></div>':'';
function akCid(){ const cs=finClients(); if(!cs.length) return null; if(isKlien()) return me().cid; if(!cs.some(c=>c.id===S.akc)) S.akc=cs[0].id; return S.akc; }
function reportNotes(cid){ return NOTES.filter(x=>x.cid===cid&&(isKlien()?!x.internal:true)); }
function notifyPic(cid,text,view){ const e=engOf(cid,'S01'); if(e&&e.pic) notify('u:'+e.pic,text,view||'akun'); }

function vAkun(){
  const cid=akCid();
  if(!cid) return head('Akuntansi klien','Laporan keuangan, proyeksi, dan analisis untuk klien.')+locked('Belum ada akses data keuangan','Managing Director dapat memberikan akses data keuangan klien kepada Anda.');
  const c=cl(cid), t=AKM.includes(S.akt)?S.akt:10, tab=AKTABS.some(x=>x[0]===S.akb)?S.akb:'laporan';
  const cs=finClients();
  const right=(isKlien()?'':'<select data-change="akc" aria-label="Klien">'+cs.map(x=>'<option value="'+x.id+'"'+(x.id===cid?' selected':'')+'>'+esc(x.nama)+'</option>').join('')+'</select>')+'<select data-change="akt" aria-label="Periode">'+AKM.map(m=>'<option value="'+m+'"'+(m===t?' selected':'')+'>'+periodLabel(m)+(isClosed(cid,m)?' (terkunci)':'')+'</option>').join('')+'</select>';
  const tabs='<div class="tabs" role="tablist" aria-label="Bagian akuntansi">'+AKTABS.map(([k,l])=>'<button role="tab" aria-selected="'+(tab===k)+'" data-a="akb" data-k="'+k+'">'+l+'</button>').join('')+'</div>';
  const fn={laporan:akLaporan,transaksi:akTransaksi,rasio:akRasio,anggaran:akAnggaran,proyeksi:akProyeksi,rekon:akRekon,review:akReview,pajak:akPajak,catatan:akCatatan,target:akTarget,periode:akPeriode,audit:akAudit}[tab];
  return head(isKlien()?'Laporan dan analisis keuangan':'Akuntansi klien',isKlien()?'Laporan keuangan usaha Anda tersusun otomatis dari catatan transaksi, lengkap dengan analisis dan catatan konsultan.':'Buku besar, laporan otomatis, proyeksi, dan analisis untuk '+esc(c.nama)+'.',right)+tabs+fn(cid,t);
}

/* ---- laporan ---- */
function akLaporan(cid,t){
  const s=buku(cid,t), c=cl(cid), cf=cashflow(cid,t), r=(n,v,b)=>'<tr><td>'+(b?'<strong>'+n+'</strong>':n)+'</td><td class="r num">'+(b?'<strong>'+rp(v)+'</strong>':rp(v))+'</td></tr>';
  const sec=n=>'<tr class="sec"><td colspan="2"><span class="eyebrow">'+n+'</span></td></tr>';
  const rows=l=>l.map(x=>r(esc(x.n),x.v)).join('');
  const neraca=tbl(['Neraca','>Saldo (Rp)'],[sec('ASET'),rows(s.asetL),rows(s.asetT),r('Total aset',s.totA,1),sec('LIABILITAS'),rows(s.liabL),rows(s.liabP),r('Total liabilitas',s.totL,1),sec('EKUITAS'),rows(s.eq),r('Laba berjalan (Jul sampai '+BLN3[t-1]+')',s.labac),r('Total ekuitas',s.totE,1),r('Total liabilitas dan ekuitas',s.totL+s.totE,1)]);
  const lr=tbl(['Laba rugi '+periodLabel(t),'>Bulan ini (Rp)'],[sec('PENDAPATAN'),rows(s.lr.pend),r('Total pendapatan',s.pend,1),sec('HARGA POKOK'),rows(s.lr.hpp),r('Laba kotor',s.pend-s.hpp,1),sec('BEBAN OPERASIONAL'),rows(s.lr.opx),r('Total beban operasional',s.opx,1),r('Laba bersih',s.laba,1)]);
  const kas=tbl(['Arus kas '+periodLabel(t),'>Jumlah (Rp)'],[r('Saldo kas awal',s.kas0,1)].concat(cf.map(g=>sec('AKTIVITAS '+g.g.toUpperCase())+(g.rows.length?g.rows.map(x=>r(esc(x.n),x.v)).join(''):'<tr><td colspan="2" class="note">Tidak ada transaksi</td></tr>')+r('Kas bersih dari aktivitas '+g.g.toLowerCase(),g.tot,1)),[r('Kenaikan (penurunan) kas',s.kas1-s.kas0,1),r('Saldo kas akhir',s.kas1,1)]));
  const labels=AKM.filter(m=>m>=T0&&m<=Math.min(t,10)).map(m=>BLN3[m-1]), ms=AKM.filter(m=>m>=T0&&m<=Math.min(t,10));
  const chart=ms.length?'<div class="card"><div class="card-h"><h2>Tren pendapatan dan laba bersih</h2></div><div class="card-b">'+bars(labels,[{name:'Pendapatan',color:'var(--brand)',vals:ms.map(m=>buku(cid,m).pend)},{name:'Laba bersih',color:'var(--copper)',vals:ms.map(m=>Math.max(0,buku(cid,m).laba))}])+'</div></div>':'';
  const noData=!TRX.some(x=>x.cid===cid&&x.t===t)?'<div class="banner info"><div><strong>Belum ada transaksi pada '+periodLabel(t)+'</strong>Laporan akan tersusun otomatis setelah transaksi dicatat pada tab Transaksi.</div></div>':'';
  return '<div class="print-area"><div class="card card-pad kop"><div><span class="eyebrow">'+esc(CONFIG.entity)+'</span><h2>Laporan Keuangan '+esc(c.nama)+'</h2><p class="note" style="margin:2px 0 0">Periode '+periodLabel(t)+' · disusun dari buku besar · data contoh untuk demo</p></div><div class="row no-print"><button class="btn" data-a="akprint">'+ic('receipt')+'Cetak atau simpan PDF</button><button class="btn" data-a="akcsv">'+ic('upload')+'Unduh Excel (CSV)</button></div></div>'+noData+
   finStrip([['mint','trend','Pendapatan',short(s.pend),periodLabel(t)],['white','wallet','Laba bersih',short(s.laba),s.pend>0?'Margin '+pc(s.laba/s.pend):'Belum ada pendapatan'],['white','filecheck','Total aset',short(s.totA),'Neraca seimbang'],['blue','wallet','Saldo kas akhir',short(s.kas1),'Kas dan bank']])+
   '<div class="grid g2e"><div class="card">'+neraca+'</div><div class="card">'+lr+'</div></div><div class="card">'+kas+'</div></div>'+chart;
}
function csvOf(cid,t){
  const s=buku(cid,t), cf=cashflow(cid,t), c=cl(cid), L=[], q=v=>'"'+String(v).replace(/"/g,'""')+'"', add=(...a)=>L.push(a.map(q).join(';'));
  add(CONFIG.entity); add('Laporan keuangan',c.nama); add('Periode',periodLabel(t)); add('Catatan','Data contoh untuk demo'); L.push('');
  add('NERACA','Saldo (Rp)'); [...s.asetL,...s.asetT].forEach(x=>add(x.n,x.v)); add('Total aset',s.totA); [...s.liabL,...s.liabP].forEach(x=>add(x.n,x.v)); add('Total liabilitas',s.totL); s.eq.forEach(x=>add(x.n,x.v)); add('Laba berjalan',s.labac); add('Total ekuitas',s.totE); L.push('');
  add('LABA RUGI','Bulan ini (Rp)'); s.lr.pend.forEach(x=>add(x.n,x.v)); s.lr.hpp.forEach(x=>add(x.n,x.v)); s.lr.opx.forEach(x=>add(x.n,x.v)); add('Laba bersih',s.laba); L.push('');
  add('ARUS KAS','Jumlah (Rp)'); add('Saldo kas awal',s.kas0); cf.forEach(g=>{ g.rows.forEach(x=>add(x.n,x.v)); add('Kas bersih aktivitas '+g.g.toLowerCase(),g.tot); }); add('Saldo kas akhir',s.kas1);
  return L.join('\r\n');
}
async function downloadText(name,text){ try{ const dl=(window.claude&&window.claude.use)?await window.claude.use('downloads'):null; if(!dl) throw new Error('na'); await dl.save({filename:name,data:'\ufeff'+text}); toast('Berkas '+name+' disimpan'); }catch(e){ if(e&&e.code==='declined') return; copyText(text,'Unduhan tidak tersedia di tampilan ini. Isi CSV disalin ke clipboard.'); } }

/* ---- transaksi dan akun ---- */
function akTransaksi(cid,t){
  const can=canInputAk(cid), f=S.trxf, A=COAOF(cid), closed=isClosed(cid,t), list=TRX.filter(x=>x.cid===cid&&x.t===t).sort((a,b)=>a.d-b.d);
  const opt=(cur)=>A.map(a=>'<option value="'+a.k+'"'+(cur===a.k?' selected':'')+'>'+a.k+' · '+esc(a.n)+'</option>').join('');
  const openMs=AKM.filter(m=>!isClosed(cid,m));
  if(can&&!openMs.includes(+f.t)) f.t=openMs[0]||10;
  const form=can?'<div class="card"><div class="card-h"><h2>Catat transaksi</h2><span class="grow"></span><span class="note">Setiap transaksi dicatat dua sisi (debit dan kredit)</span></div><div class="card-b"><div class="formgrid">'+
   '<label class="f">Jenis transaksi<select data-change="akf" data-f="trxf" data-k="jenis">'+JENIS.map(j=>'<option value="'+j[0]+'"'+(f.jenis===j[0]?' selected':'')+'>'+j[1]+'</option>').join('')+'</select></label>'+
   '<label class="f">Bulan<select data-change="akf" data-f="trxf" data-k="t">'+openMs.map(m=>'<option value="'+m+'"'+(+f.t===m?' selected':'')+'>'+periodLabel(m)+'</option>').join('')+'</select></label>'+
   '<label class="f">Tanggal (1 sampai 28)<input type="number" min="1" max="28" data-input="akf" data-f="trxf" data-k="d" value="'+esc(f.d)+'"></label>'+
   '<label class="f">Jumlah (Rp)<input type="number" min="1" inputmode="numeric" data-input="akf" data-f="trxf" data-k="jml" value="'+esc(f.jml)+'"></label>'+
   '<label class="f">Akun debit<select data-change="akf" data-f="trxf" data-k="deb">'+opt(f.deb)+'</select></label>'+
   '<label class="f">Akun kredit<select data-change="akf" data-f="trxf" data-k="kre">'+opt(f.kre)+'</select></label>'+
   '<label class="f" style="grid-column:1/-1">Keterangan<input type="text" data-input="akf" data-f="trxf" data-k="ket" value="'+esc(f.ket)+'" placeholder="Contoh: Penjualan toko minggu ke-2"></label></div>'+
   '<div class="row" style="margin-top:12px"><button class="btn primary" data-a="trxadd">'+ic('plus')+'Catat transaksi</button></div></div></div>':'<div class="banner info"><div><strong>Mode lihat saja</strong>Pencatatan transaksi dilakukan oleh klien dan tim akuntansi.</div></div>';
  const trows=list.map(x=>'<tr><td class="num">'+dayLbl(x.t,x.d)+'</td><td>'+esc(x.ket)+'<div class="note">'+esc(x.by)+'</div></td><td>'+x.deb+' · '+esc((accOf(cid,x.deb)||{n:'-'}).n)+'</td><td>'+x.kre+' · '+esc((accOf(cid,x.kre)||{n:'-'}).n)+'</td><td class="r num">'+rp(x.jml)+'</td><td class="r">'+(can&&!closed?'<button class="btn sm" data-a="trxdel" data-id="'+x.id+'" aria-label="Hapus transaksi">'+ic('x')+'</button>':ic('lock'))+'</td></tr>');
  const s=buku(cid,t);
  const accTable=tbl(['Kode','Akun','Jenis','>Saldo ('+periodLabel(t)+')'],s.accs.map(a=>'<tr><td class="mono">'+a.k+'</td><td>'+esc(a.n)+'</td><td>'+TIPE_AKUN[a.t]+((ACCX[cid]||[]).some(x=>x.k===a.k)?' '+pill('info','Khusus klien'):'')+'</td><td class="r num">'+rp(a.v)+'</td></tr>'));
  const accForm=canStaffAk(cid)?'<div class="card-b" style="border-top:1px solid var(--line)"><div class="formgrid"><label class="f">Kode akun<input type="text" maxlength="4" data-input="akf" data-f="accf" data-k="k" value="'+esc(S.accf.k)+'" placeholder="6400"></label><label class="f">Nama akun<input type="text" data-input="akf" data-f="accf" data-k="n" value="'+esc(S.accf.n)+'" placeholder="Beban Pemasaran"></label><label class="f">Jenis<select data-change="akf" data-f="accf" data-k="t"><option value="beban"'+(S.accf.t==='beban'?' selected':'')+'>Beban</option><option value="pend"'+(S.accf.t==='pend'?' selected':'')+'>Pendapatan</option><option value="aset"'+(S.accf.t==='aset'?' selected':'')+'>Aset lancar</option><option value="liab"'+(S.accf.t==='liab'?' selected':'')+'>Liabilitas lancar</option></select></label></div><div class="row" style="margin-top:10px"><button class="btn" data-a="accadd">'+ic('plus')+'Tambah akun untuk klien ini</button></div></div>':'';
  return lockBanner(cid,t)+form+'<div class="card"><div class="card-h"><h2>Transaksi '+periodLabel(t)+'</h2><span class="grow"></span><span class="note">'+list.length+' transaksi</span></div>'+(list.length?tbl(['Tanggal','Keterangan','Debit','Kredit','>Jumlah',''],trows):emptyBox('Belum ada transaksi pada periode ini.'))+'</div><div class="card"><div class="card-h"><h2>Daftar akun</h2><span class="grow"></span><span class="note">Akun standar dapat ditambah sesuai kebutuhan klien</span></div>'+accTable+accForm+'</div>';
}

/* ---- rasio ---- */
const ratioCard=r=>'<div class="card card-pad rcard '+r.st+'"><div style="display:flex;justify-content:space-between;gap:8px;align-items:center"><span class="eyebrow">'+r.label.toUpperCase()+'</span>'+pill(r.st==='mute'?'mute':r.st,RST[r.st])+'</div><div class="rval">'+r.val+'</div><div class="note">'+esc(r.hint)+'</div></div>';
function akRasio(cid,t){
  const rs=ratios(cid,t), ms=AKM.filter(m=>m<=Math.min(t,10)&&m>=T0);
  const hist=tbl(['Periode','>Margin laba','>Rasio lancar','>Perputaran piutang','>Umur kas'],ms.map(m=>{ const x=ratios(cid,m); return '<tr><td>'+periodLabel(m)+'</td>'+x.map(y=>'<td class="r num">'+y.val+'</td>').join('')+'</tr>'; }));
  return '<div class="grid g2e">'+rs.map(ratioCard).join('')+'</div><div class="card"><div class="card-h"><h2>Perkembangan rasio</h2></div>'+hist+'</div><p class="note">Ambang hijau, kuning, dan merah adalah pedoman umum untuk usaha kecil dan menengah. Konsultan dapat menyesuaikannya dengan karakter usaha klien.</p>';
}

/* ---- anggaran ---- */
function akAnggaran(cid,t){
  const rows=budRows(cid,t), can=canStaffAk(cid), warn=rows.filter(r=>r.st==='bad');
  const body=rows.map(r=>'<tr><td><strong>'+esc(r.n)+'</strong></td><td class="r">'+(can?'<input type="number" min="0" step="1000" data-change="bud" data-k="'+r.k+'" value="'+r.bud+'" style="width:140px;text-align:right" aria-label="Anggaran '+esc(r.n)+'">':'<span class="num">'+rp(r.bud)+'</span>')+'</td><td class="r num">'+rp(r.real)+'</td><td class="r num">'+(r.sel>0?'+':'')+rp(r.sel)+'</td><td class="r num">'+(r.bud?(r.pct>0?'+':'')+pc(r.pct):'–')+'</td><td>'+pill(r.st==='mute'?'mute':r.st,r.st==='good'?'Sesuai':r.st==='warn'?'Mendekati batas':r.st==='bad'?(r.inc?'Di bawah target':'Melebihi anggaran'):'Belum ada anggaran')+'</td></tr>').join('');
  return (warn.length?'<div class="banner bad"><div><strong>'+warn.length+' pos menyimpang lebih dari 10% pada '+periodLabel(t)+'</strong>'+warn.map(r=>esc(r.n)).join(', ')+'. Periksa penyebabnya bersama klien.</div></div>':'')+
   '<div class="card"><div class="card-h"><h2>Anggaran dan realisasi '+periodLabel(t)+'</h2><span class="grow"></span><span class="note">Beban dianggap menyimpang bila lebih dari 10% di atas anggaran</span></div>'+tbl(['Pos','>Anggaran bulanan','>Realisasi','>Selisih','>Selisih %','Status'],[body])+'</div>'+(can?'':'<p class="note">Anggaran diatur oleh konsultan MandaGroup bersama Anda.</p>');
}

/* ---- proyeksi ---- */
function lineChart(labels,series){
  const W=640,H=230,pl=70,pr=14,pt=14,pb=30, all=series.flatMap(x=>x.vals); let mn=Math.min(0,...all), mx=Math.max(1,...all); if(mx===mn) mx=mn+1;
  const X=i=>pl+(labels.length<2?0:i*(W-pl-pr)/(labels.length-1)), Y=v=>pt+(mx-v)*(H-pt-pb)/(mx-mn); let g='';
  for(let i=0;i<=4;i++){ const v=mn+(mx-mn)*i/4, y=Y(v); g+='<line x1="'+pl+'" x2="'+(W-pr)+'" y1="'+y+'" y2="'+y+'" style="stroke:var(--line)"/><text x="'+(pl-8)+'" y="'+(y+4)+'" text-anchor="end" style="fill:var(--muted);font-size:11px">'+sh(v)+'</text>'; }
  const step=Math.max(1,Math.ceil(labels.length/8)); labels.forEach((l,i)=>{ if(i%step===0||i===labels.length-1) g+='<text x="'+X(i)+'" y="'+(H-8)+'" text-anchor="middle" style="fill:var(--muted);font-size:11px">'+l+'</text>'; });
  series.forEach(x=>{ g+='<polyline fill="none" style="stroke:'+x.color+';stroke-width:2.5" points="'+x.vals.map((v,i)=>X(i)+','+Y(v)).join(' ')+'"/>'+x.vals.map((v,i)=>'<circle cx="'+X(i)+'" cy="'+Y(v)+'" r="3" style="fill:'+x.color+'"/>').join(''); });
  return '<svg viewBox="0 0 '+W+' '+H+'" class="lc" role="img" aria-label="Grafik proyeksi: '+series.map(x=>x.name).join(', ')+'">'+g+'</svg><div class="legend">'+series.map(x=>'<span><i style="background:'+x.color+'"></i>'+x.name+'</span>').join('')+'</div>';
}
function akProyeksi(cid){
  const a=pa(cid), can=canStaffAk(cid), hor=+S.prhor||6, met=S.prmet, runs=SCN.map(sc=>({sc,r:projRun(cid,hor,sc)})), cur=runs.find(x=>x.sc.k===S.prsc)||runs[1];
  const mlabel={laba:'Laba bersih bulanan',kas:'Saldo kas akhir',rev:'Pendapatan bulanan'}[met]||'Laba bersih bulanan', mkey=met==='kas'?'kas':met==='rev'?'rev':'laba';
  const asm=(l,k,v,u,help)=>'<label class="f">'+l+(u?' ('+u+')':'')+'<input type="number" step="'+(k==='fx'?'100000':'0.1')+'" data-change="asm" data-k="'+k+'" value="'+v+'"'+(can?'':' disabled')+'><span class="note">'+help+'</span></label>';
  const base=buku(cid,BASEM);
  const summ=tbl(['Skenario','>Pendapatan bulan ke-'+hor,'>Laba bulan ke-'+hor,'>Total laba '+hor+' bulan','>Saldo kas akhir'],runs.map(x=>{ const l=x.r[x.r.length-1]; return '<tr><td><strong>'+x.sc.n+'</strong><div class="note">Pertumbuhan '+(a.g+x.sc.dg).toFixed(1).replace('.',',')+'% per bulan</div></td><td class="r num">'+rp(l.rev)+'</td><td class="r num">'+rp(l.laba)+'</td><td class="r num">'+rp(x.r.reduce((s,y)=>s+y.laba,0))+'</td><td class="r num">'+rp(l.kas)+'</td></tr>'; }));
  const det=tbl(['Bulan','>Pendapatan','>Beban','>Laba bersih','>Saldo kas'],cur.r.map(y=>'<tr><td>'+y.lbl+'</td><td class="r num">'+rp(y.rev)+'</td><td class="r num">'+rp(y.beban)+'</td><td class="r num">'+rp(y.laba)+'</td><td class="r num">'+rp(y.kas)+'</td></tr>'));
  return '<div class="card"><div class="card-h"><h2>Asumsi proyeksi</h2><span class="grow"></span><span class="note">Dasar: realisasi '+periodLabel(BASEM)+' · pendapatan '+short(base.pend)+'</span></div><div class="card-b"><div class="formgrid">'+asm('Pertumbuhan pendapatan dasar','g',a.g,'% per bulan','Skenario lain bergeser 2 poin')+asm('Beban variabel','vp',a.vp,'% pendapatan','HPP dan beban operasional di luar gaji dan sewa')+asm('Beban tetap bulanan','fx',Math.round(a.fx),'Rp','Gaji dan sewa')+'</div>'+(can?'':'<p class="note" style="margin:10px 0 0">Asumsi diatur oleh konsultan MandaGroup.</p>')+'</div></div>'+
   '<div class="card"><div class="card-h"><h2>'+mlabel+'</h2><span class="grow"></span><div class="tabs" role="tablist">'+[3,6,12].map(h=>'<button role="tab" aria-selected="'+(hor===h)+'" data-a="prhor" data-k="'+h+'">'+h+' bulan</button>').join('')+'</div><div class="tabs" role="tablist">'+[['laba','Laba'],['kas','Kas'],['rev','Pendapatan']].map(([k,l])=>'<button role="tab" aria-selected="'+(met===k||(k==='laba'&&!['kas','rev'].includes(met)))+'" data-a="prmet" data-k="'+k+'">'+l+'</button>').join('')+'</div></div><div class="card-b">'+lineChart(cur.r.map(y=>y.lbl),runs.map(x=>({name:x.sc.n,color:x.sc.color,vals:x.r.map(y=>y[mkey])})))+'</div></div>'+
   '<div class="card"><div class="card-h"><h2>Perbandingan skenario</h2></div>'+summ+'</div>'+
   '<div class="card"><div class="card-h"><h2>Rincian skenario</h2><span class="grow"></span><div class="tabs" role="tablist">'+SCN.map(x=>'<button role="tab" aria-selected="'+(cur.sc.k===x.k)+'" data-a="prsc" data-k="'+x.k+'">'+x.n+'</button>').join('')+'</div></div>'+det+'</div>'+
   '<p class="note">Proyeksi bersifat perkiraan. Perubahan kas dihitung sama dengan laba bersih, tanpa memperhitungkan modal kerja, pembelian aset, dan angsuran. Hasilnya bukan jaminan kinerja usaha.</p>';
}

/* ---- rekonsiliasi bank ---- */
function akRekon(cid,t){
  const s=buku(cid,t), ex=BANKX.filter(x=>x.cid===cid&&x.t===t), net=ex.reduce((a,x)=>a+x.jml,0), bank=s.kas1+net, matched=TRX.filter(x=>x.cid===cid&&x.t===t&&(x.deb==='1100'||x.kre==='1100')).length, closed=isClosed(cid,t), can=canInputAk(cid);
  const rows=ex.map(x=>'<tr><td class="num">'+dayLbl(x.t,x.d)+'</td><td>'+esc(x.ket)+'</td><td class="r num">'+(x.jml>0?'+':'')+rp(x.jml)+'</td><td>'+(x.st==='diselidiki'?pill('warn','Sedang diselidiki'):pill('bad','Belum ada di buku'))+'</td><td class="r">'+(can&&!closed?'<button class="btn sm" data-a="bkcatat" data-id="'+x.id+'">Catat ke buku</button>':'')+(canReviewAk(cid)&&!closed&&x.st!=='diselidiki'?' <button class="btn sm" data-a="bkinv" data-id="'+x.id+'">Selidiki</button>':'')+'</td></tr>');
  return lockBanner(cid,t)+finStrip([['white','wallet','Saldo menurut buku',short(s.kas1),'Akun 1100 Kas dan Bank'],['white','wallet','Saldo menurut bank',short(bank),'Rekening koran '+periodLabel(t)],[net===0?'mint':'blue',net===0?'filecheck':'clock','Selisih',net===0?'Rp 0':rp(net),net===0?'Sudah cocok':ex.length+' item perlu ditindaklanjuti'],['white','check','Transaksi cocok',String(matched),'Transaksi kas yang sesuai dengan mutasi bank']])+
   '<div class="card"><div class="card-h"><h2>Item mutasi bank yang belum cocok</h2></div>'+(rows.length?tbl(['Tanggal','Keterangan mutasi','>Jumlah','Status',''],rows):emptyBox('Seluruh mutasi bank pada '+periodLabel(t)+' sudah sesuai dengan buku.'))+'</div><p class="note">Item yang sudah dicatat ke buku otomatis hilang dari daftar ini. Mutasi tanpa keterangan sebaiknya diselidiki bersama klien sebelum dicatat.</p>';
}

/* ---- review per akun ---- */
function akReview(cid,t){
  const s=buku(cid,t), fl=ACCFIND.filter(x=>x.cid===cid), can=canReviewAk(cid), k=S.rvk;
  const tb=tbl(['Kode','Akun','>Saldo','Temuan terbuka'],s.accs.map(a=>{ const o=fl.filter(x=>x.k===a.k&&x.st!=='ditutup').length; return '<tr><td class="mono">'+a.k+'</td><td>'+esc(a.n)+'</td><td class="r num">'+rp(a.v)+'</td><td>'+(o?pill('warn',o+' temuan'):'<span class="note">–</span>')+'</td></tr>'; }));
  const form=can?'<div class="card"><div class="card-h"><h2>Beri temuan pada akun</h2></div><div class="card-b"><div class="formgrid"><label class="f">Akun<select data-change="akf" data-f="afs" data-k="k">'+COAOF(cid).map(a=>'<option value="'+a.k+'"'+(S.afs.k===a.k?' selected':'')+'>'+a.k+' · '+esc(a.n)+'</option>').join('')+'</select></label><label class="f">Tingkat<select data-change="akf" data-f="afs" data-k="sev"><option value="tinggi"'+(S.afs.sev==='tinggi'?' selected':'')+'>Tinggi</option><option value="sedang"'+(S.afs.sev==='sedang'?' selected':'')+'>Sedang</option><option value="rendah"'+(S.afs.sev==='rendah'?' selected':'')+'>Rendah</option></select></label></div><label class="f" style="margin-top:10px">Uraian temuan dan saran perbaikan<textarea data-input="akf" data-f="afs" data-k="teks" placeholder="Tulis temuan yang jelas dan dapat ditindaklanjuti klien.">'+esc(S.afs.teks)+'</textarea></label><div class="row" style="margin-top:10px"><button class="btn primary" data-a="afadd">'+ic('plus')+'Tambah temuan</button></div></div></div>':'';
  const list=fl.length?fl.map(f=>{ const a=accOf(cid,f.k)||{n:f.k}; let act='';
    if(isKlien()&&f.st==='terbuka') act='<label class="f">Tanggapan Anda<textarea data-input="afr" data-id="'+f.id+'" placeholder="Jelaskan perbaikan atau berikan klarifikasi.">'+esc(S.afr[f.id]||'')+'</textarea></label><div class="row"><button class="btn primary sm" data-a="afresp" data-id="'+f.id+'">Kirim tanggapan</button></div>';
    if(can&&f.st==='ditanggapi') act='<div class="row"><button class="btn primary sm" data-a="afclose" data-id="'+f.id+'">Terima dan tutup temuan</button></div>';
    return '<div class="card-b" style="border-top:1px solid var(--line);display:flex;flex-direction:column;gap:8px"><div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap"><strong>'+f.k+' · '+esc(a.n)+'</strong>'+pill(f.sev==='tinggi'?'bad':f.sev==='sedang'?'warn':'mute',f.sev[0].toUpperCase()+f.sev.slice(1))+pill(f.st==='ditutup'?'good':f.st==='ditanggapi'?'info':'warn',f.st==='ditutup'?'Ditutup':f.st==='ditanggapi'?'Menunggu persetujuan':'Perlu ditanggapi')+'<span class="grow"></span><small class="note">'+esc(f.by)+' · '+esc(f.ts)+'</small></div><div>'+esc(f.teks)+'</div>'+(f.resp?'<div class="banner info" style="padding:8px 12px"><div><strong>Tanggapan klien</strong>'+esc(f.resp)+'</div></div>':'')+act+'</div>'; }).join(''):emptyBox('Belum ada temuan pada akun.');
  return '<div class="card"><div class="card-h"><h2>Saldo akun '+periodLabel(t)+'</h2></div>'+tb+'</div>'+form+'<div class="card"><div class="card-h"><h2>Temuan per akun</h2></div>'+list+'</div>';
}

/* ---- pajak ---- */
function akPajak(cid,t){
  const x=taxEst(cid,t), can=canStaffAk(cid), s=buku(cid,t);
  const ctl=can?'<div class="card-b" style="border-top:1px solid var(--line);display:flex;gap:16px;flex-wrap:wrap;align-items:end"><label class="f">Skema pajak penghasilan<select data-change="taxm"><option value="final"'+(x.mode==='final'?' selected':'')+'>PPh Final UMKM 0,5%</option><option value="badan"'+(x.mode==='badan'?' selected':'')+'>PPh Badan 22%</option></select></label><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" data-change="taxpkp"'+(x.pkp?' checked':'')+'> Pengusaha Kena Pajak (PKP), kenakan PPN</label></div>':'';
  const ppn=x.ppn?tbl(['PPN ('+(CONFIG.ppn*100)+'%)','>Bulan ini','>Kumulatif sejak Juli'],['<tr><td>PPN keluaran (dari pendapatan)</td><td class="r num">'+rp(x.ppn.keluaran)+'</td><td class="r num">'+rp(x.ppn.kKel)+'</td></tr>','<tr><td>PPN masukan (dari pembelian dan HPP)</td><td class="r num">'+rp(x.ppn.masukan)+'</td><td class="r num">'+rp(x.ppn.kMas)+'</td></tr>','<tr><td><strong>Estimasi PPN kurang bayar</strong></td><td class="r num"><strong>'+rp(x.ppn.keluaran-x.ppn.masukan)+'</strong></td><td class="r num"><strong>'+rp(x.ppn.kKel-x.ppn.kMas)+'</strong></td></tr>']):'<div class="card-b note">Usaha ini tidak dikukuhkan sebagai PKP, sehingga PPN tidak dihitung.</div>';
  return '<div class="banner info"><div><strong>Angka di bawah adalah estimasi</strong>Dihitung otomatis dari transaksi, belum memperhitungkan koreksi fiskal, fasilitas, dan masa berlaku tarif. Tim perpajakan MandaGroup memverifikasi sebelum Anda melapor.</div></div>'+
   finStrip([['white','trend','Pendapatan '+BLN3[t-1],short(s.pend),'Dasar perhitungan'],['mint','receipt','Estimasi PPh bulan ini',short(x.pph.bulan),x.pph.nama],['white','wallet','Estimasi PPh kumulatif',short(x.pph.kum),'Sejak Juli 2026 (data demo)'],['blue','receipt','Estimasi PPN kurang bayar',x.ppn?short(x.ppn.keluaran-x.ppn.masukan):'–',x.ppn?'Bulan ini':'Bukan PKP']])+
   '<div class="card"><div class="card-h"><h2>Pajak penghasilan</h2></div>'+tbl(['Komponen','>Bulan ini','>Kumulatif sejak Juli'],['<tr><td>'+esc(x.pph.nama)+'</td><td class="r num">'+rp(x.pph.bulan)+'</td><td class="r num">'+rp(x.pph.kum)+'</td></tr>'])+ctl+'</div><div class="card"><div class="card-h"><h2>Pajak pertambahan nilai</h2></div>'+ppn+'</div>';
}

/* ---- catatan konsultan ---- */
const notePill=st=>pill(st==='selesai'?'good':st==='dalam proses'?'info':'warn',st==='selesai'?'Selesai':st==='dalam proses'?'Dalam proses':'Baru');
function catatanCard(c){ const n=NOTES.filter(x=>x.cid===c.id&&!x.internal&&x.st!=='selesai').slice(0,3); if(!n.length) return ''; return '<section class="card"><div class="card-h"><h2>Catatan konsultan</h2><span class="grow"></span><button class="link" data-a="akgo" data-k="catatan">Lihat semua</button></div><div class="list">'+n.map(x=>'<div class="item"><div class="grow"><strong>'+esc(x.judul)+'</strong><small>'+esc(x.by)+' · '+esc(x.ts)+'</small></div>'+notePill(x.st)+'</div>').join('')+'</div></section>'; }
function akCatatan(cid){
  const ns=reportNotes(cid), can=canReviewAk(cid);
  const form=can?'<div class="card"><div class="card-h"><h2>Tambah catatan atau rekomendasi</h2></div><div class="card-b" style="display:flex;flex-direction:column;gap:10px"><label class="f">Judul<input type="text" data-input="akf" data-f="notef" data-k="judul" value="'+esc(S.notef.judul)+'"></label><label class="f">Isi rekomendasi<textarea data-input="akf" data-f="notef" data-k="teks">'+esc(S.notef.teks)+'</textarea></label><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" data-change="akf" data-f="notef" data-k="internal"'+(S.notef.internal?' checked':'')+'> Catatan internal (tidak terlihat klien)</label><div class="row"><button class="btn primary" data-a="noteadd">'+ic('plus')+'Simpan catatan</button></div></div></div>':'';
  const list=ns.length?ns.map(x=>'<div class="card-b" style="border-top:1px solid var(--line);display:flex;flex-direction:column;gap:6px"><div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap"><strong>'+esc(x.judul)+'</strong>'+notePill(x.st)+(x.internal?pill('mute','Internal'):'')+'<span class="grow"></span><small class="note">'+esc(x.by)+' · '+esc(x.ts)+'</small></div><div>'+esc(x.teks)+'</div><div class="row">'+(x.st==='baru'?'<button class="btn sm" data-a="notest" data-id="'+x.id+'" data-s="dalam proses">Mulai tindak lanjut</button>':'')+(x.st!=='selesai'?'<button class="btn sm" data-a="notest" data-id="'+x.id+'" data-s="selesai">Tandai selesai</button>':'')+'</div></div>').join(''):emptyBox('Belum ada catatan dari konsultan.');
  return form+'<div class="card"><div class="card-h"><h2>Catatan dan rekomendasi konsultan</h2></div>'+list+'</div>';
}

/* ---- target dan jadwal ---- */
const METRIK={laba:'Laba bersih bulanan',pendapatan:'Pendapatan bulanan',kas:'Saldo kas akhir'};
function akTarget(cid){
  const s=buku(cid,10), val={laba:s.laba,pendapatan:s.pend,kas:s.kas1}, can=isKlien()||canStaffAk(cid), gs=TARGETS.filter(x=>x.cid===cid), ss=SESI.filter(x=>x.cid===cid).sort((a,b)=>a.tgl-b.tgl), staff=canReviewAk(cid);
  const gl=gs.length?gs.map(g=>{ const v=Math.max(0,val[g.metrik]||0), p=g.nilai>0?Math.min(1,v/g.nilai):0; return '<div class="card-b" style="border-top:1px solid var(--line)"><div style="display:flex;justify-content:space-between;gap:8px;align-items:center"><strong>'+METRIK[g.metrik]+'</strong><span class="num">'+rp(v)+' dari '+rp(g.nilai)+'</span></div><div class="meter" role="img" aria-label="Progres '+Math.round(p*100)+' persen"><i style="width:'+Math.round(p*100)+'%"></i></div><div style="display:flex;justify-content:space-between;align-items:center"><small class="note">Capaian '+pc(g.nilai>0?v/g.nilai:0)+' (data '+periodLabel(10)+')</small>'+(can?'<button class="btn sm" data-a="tgdel" data-id="'+g.id+'">Hapus</button>':'')+'</div></div>'; }).join(''):emptyBox('Belum ada target. Tambahkan target agar progresnya dipantau tiap bulan.');
  const gform=can?'<div class="card-b" style="border-top:1px solid var(--line)"><div class="formgrid"><label class="f">Metrik<select data-change="akf" data-f="tgf" data-k="metrik">'+Object.entries(METRIK).map(([k,l])=>'<option value="'+k+'"'+(S.tgf.metrik===k?' selected':'')+'>'+l+'</option>').join('')+'</select></label><label class="f">Target (Rp)<input type="number" min="1" data-input="akf" data-f="tgf" data-k="nilai" value="'+esc(S.tgf.nilai)+'"></label></div><div class="row" style="margin-top:10px"><button class="btn" data-a="tgadd">'+ic('plus')+'Tambah target</button></div></div>':'';
  const sl=ss.length?ss.map(x=>'<div class="item"><div class="grow"><strong>'+esc(x.topik)+'</strong><small>'+dfmt(x.tgl)+(x.pic?' · '+esc(x.pic):'')+'</small></div>'+pill(x.st==='dijadwalkan'?'good':x.st==='diajukan'?'warn':'mute',x.st==='dijadwalkan'?'Dijadwalkan':x.st==='diajukan'?'Menunggu konfirmasi':'Dibatalkan')+(staff&&x.st==='diajukan'?'<button class="btn sm primary" data-a="sesok" data-id="'+x.id+'">Konfirmasi</button>':'')+(x.st!=='batal'&&(staff||isKlien())?'<button class="btn sm" data-a="sesbatal" data-id="'+x.id+'">Batalkan</button>':'')+'</div>').join(''):emptyBox('Belum ada sesi konsultasi.');
  const sform=(isKlien()||canStaffAk(cid))?'<div class="card-b" style="border-top:1px solid var(--line)"><div class="formgrid"><label class="f">Tanggal<input type="date" data-input="akf" data-f="sesf" data-k="tgl" value="'+esc(S.sesf.tgl)+'"></label><label class="f">Topik<input type="text" data-input="akf" data-f="sesf" data-k="topik" value="'+esc(S.sesf.topik)+'" placeholder="Contoh: Evaluasi arus kas"></label></div><div class="row" style="margin-top:10px"><button class="btn" data-a="sesadd">'+ic('plus')+(isKlien()?'Ajukan sesi konsultasi':'Jadwalkan sesi')+'</button></div></div>':'';
  return '<div class="grid g2e"><div class="card"><div class="card-h"><h2>Target keuangan</h2></div>'+gl+gform+'</div><div class="card"><div class="card-h"><h2>Jadwal konsultasi</h2></div><div class="list">'+sl+'</div>'+sform+'</div></div>';
}

/* ---- tutup buku ---- */
function akPeriode(cid){
  const can=isSU()||(isTim()&&me().div==='review'), rows=AKM.map(m=>{ const sb=SUBS.find(x=>x.cid===cid&&x.t===m), cl_=CLOSED[cid+'|'+m], ok=sb&&sb.status==='disetujui';
    return '<tr><td><strong>'+periodLabel(m)+'</strong></td><td>'+(sb?subPill(sb,isKlien()):'<span class="note">Belum ada laporan</span>')+'</td><td>'+(cl_?pill('good','Terkunci')+'<div class="note">'+esc(cl_.by)+' · '+esc(cl_.ts)+'</div>':pill('mute','Terbuka'))+'</td><td class="r">'+(cl_?(isSU()?'<button class="btn sm" data-a="akreopen" data-t="'+m+'">Buka kembali</button>':''):(can?'<button class="btn sm primary" data-a="akclose" data-t="'+m+'"'+(ok?'':' disabled title="Periode dapat ditutup setelah laporan disetujui"')+'>Tutup buku</button>':''))+'</td></tr>'; });
  return '<div class="banner info"><div><strong>Tutup buku bulanan</strong>Setelah laporan disetujui reviewer, periode dikunci agar angka yang sudah direview tidak berubah. Pembukaan kembali hanya oleh Managing Director dan tercatat di jejak audit.</div></div><div class="card"><div class="card-h"><h2>Status periode</h2></div>'+tbl(['Periode','Status review','Status buku','>Tindakan'],rows)+'</div>';
}

/* ---- jejak audit ---- */
const AKAT=['Transaksi','Akun','Anggaran','Proyeksi','Review','Rekonsiliasi','Periode','Pajak','Catatan','Target','Laporan'];
function akAudit(cid){
  const q=(S.afq||'').toLowerCase(), list=AUDIT.filter(x=>x.cid===cid&&(S.afk==='semua'||x.kat===S.afk)&&(!q||(x.text+x.who).toLowerCase().includes(q)));
  return '<div class="card"><div class="card-h"><h2>Jejak audit</h2><span class="grow"></span><select data-change="afk" aria-label="Kategori"><option value="semua">Semua kategori</option>'+AKAT.map(k=>'<option'+(S.afk===k?' selected':'')+'>'+k+'</option>').join('')+'</select><input type="search" data-change="afq" value="'+esc(S.afq)+'" placeholder="Cari pengguna atau uraian" aria-label="Cari" style="width:200px"></div>'+(list.length?tbl(['Waktu','Pengguna','Kategori','Uraian'],list.slice(0,200).map(x=>'<tr><td class="num">'+esc(x.ts)+'</td><td>'+esc(x.who)+'</td><td>'+pill('mute',esc(x.kat))+'</td><td>'+esc(x.text)+'</td></tr>')):emptyBox('Tidak ada catatan yang cocok dengan filter.'))+'</div><p class="note">Setiap perubahan angka, anggaran, asumsi, dan status periode dicatat otomatis beserta pengguna dan waktunya.</p>';
}

/* ---------- paket: tampilan ---------- */
const chipsOf=c=>'<div class="chips">'+pkSvcs(c).map(id=>'<span class="chip">'+SVSHORT[id]+'</span>').join('')+addons(c).map(e=>'<span class="chip add">+ '+SVSHORT[e.sid]+'</span>').join('')+'</div>';
function paketCard(c){ const p=pk(c.paket), nx=c.paketNext;
  return '<section class="card card-pad paket-card"><div><span class="eyebrow">PAKET LAYANAN ANDA</span><h2>Paket '+esc(p.nama)+'</h2><p class="note" style="margin:4px 0 0">'+esc(p.tag)+'</p><div class="pk-price">'+rp(p.harga)+'<small> / bulan</small></div></div><div style="display:flex;flex-direction:column;gap:12px"><div><div class="note" style="margin-bottom:6px">Layanan yang Anda gunakan</div>'+chipsOf(c)+'</div>'+(nx?'<div class="banner info" style="padding:10px 12px"><div><strong>Perubahan ke paket '+esc(pk(nx.id).nama)+' berlaku '+dfmt(nx.berlaku)+'</strong></div></div>':'')+'<button class="link" style="text-align:left" data-a="go" data-v="layanan">Lihat detail dan bandingkan paket</button></div></section>'; }
function pkCard(p,cur,btn){ const sv=paketSaving(p); return '<div class="pk'+(cur?' cur':'')+'"><div style="display:flex;justify-content:space-between;align-items:center;gap:8px"><strong style="font:700 17px var(--f-display)">'+esc(p.nama)+'</strong>'+(cur?pill('good','Paket Anda'):'')+'</div><div class="note">'+esc(p.tag)+'</div><div class="pk-price">'+rp(p.harga)+'<small> / bulan</small></div>'+(sv>0?'<div class="note" style="color:var(--good)">Hemat '+rp(sv)+' dibanding layanan satuan</div>':'<div class="note">&nbsp;</div>')+'<div class="chips">'+p.svc.map(id=>'<span class="chip">'+SVSHORT[id]+'</span>').join('')+'</div><ul>'+p.fitur.map(f=>'<li>'+esc(f)+'</li>').join('')+'</ul><div class="grow"></div>'+(btn||'')+'</div>'; }
function vLayananKlien(){
  const c=cl(me().cid), p=pk(c.paket), ad=addons(c), nx=c.paketNext, mine=ENG.filter(e=>e.cid===c.id).sort((a,b)=>(pkSvcs(c).includes(b.sid)?1:0)-(pkSvcs(c).includes(a.sid)?1:0)), other=SERVICES.filter(x=>!mine.some(e=>e.sid===x.id));
  return head('Layanan dan paket saya','Paket layanan akuntansi bulanan dan layanan '+CONFIG.brand+' yang aktif untuk '+esc(c.nama)+'.')+
  (nx?'<div class="banner info"><div><strong>Perubahan paket dijadwalkan: '+esc(p.nama)+' → '+esc(pk(nx.id).nama)+'</strong>Berlaku '+dfmt(nx.berlaku)+'. Tagihan bulan berikutnya mengikuti paket baru.</div></div>':'')+
  '<div class="grid g4 dashboard-stats">'+stat('mint','filecheck','Paket aktif',esc(p.nama),rp(p.harga)+' per bulan')+stat('white','list','Layanan aktif',String(mine.length),ad.length?pkSvcs(c).length+' dalam paket + '+ad.length+' tambahan':'Semua dalam paket')+stat('white','clock','Hasil review',svc('S01').sla+' hari kerja','Sejak dokumen lengkap diterima')+stat('blue','receipt','Klien sejak',c.sejak,'Tagihan setiap awal bulan')+'</div>'+
  '<div class="card"><div class="card-h"><h2>Layanan yang Anda gunakan</h2></div><div class="tbl-wrap"><table><thead><tr><th>Layanan</th><th>Siklus</th><th>Penanggung jawab</th><th>Sumber</th></tr></thead><tbody>'+mine.map(e=>'<tr><td><strong>'+esc(svc(e.sid).nama)+'</strong><div class="note">'+esc(svc(e.sid).desc)+'</div></td><td>'+svc(e.sid).siklus+'</td><td>'+esc(UNAME(e.pic))+'<div class="note">'+DIV[svc(e.sid).div]+'</div></td><td>'+(pkSvcs(c).includes(e.sid)?pill('good','Termasuk paket'):pill('info','Layanan tambahan'))+'</td></tr>').join('')+'</tbody></table></div></div>'+
  '<div><span class="eyebrow">PILIHAN PAKET</span><h2 class="sub-title" style="margin-top:6px">Bandingkan paket layanan bulanan</h2><div class="pk-grid">'+PAKET.map(x=>pkCard(x,x.id===c.paket,x.id===c.paket?'':'<button class="btn sm" data-a="reqpaket" data-id="'+x.id+'">Ajukan perubahan ke '+esc(x.nama)+'</button>')).join('')+'</div><p class="note" style="margin-top:10px">Perubahan paket diproses tim '+CONFIG.brand+' dan berlaku mulai periode berikutnya. Harga sudah di luar PPN.</p></div>'+
  (other.length?'<h2 class="sub-title">Layanan tambahan yang tersedia</h2><div class="grid g3">'+other.map(x=>'<div class="card card-b" style="display:flex;flex-direction:column;gap:8px"><strong>'+esc(x.nama)+'</strong><span class="note">'+esc(x.desc)+'</span><span class="num">Mulai '+rp(x.harga)+' · '+x.siklus+'</span><div class="grow"></div><button class="btn sm" data-a="reqsvc" data-id="'+x.id+'">Ajukan layanan</button></div>').join('')+'</div>':'');
}
const canPaket=()=>isSU()||isTim();
function vKlienPaket(c){
  const p=pk(c.paket), nx=c.paketNext, ad=addons(c), log=PAKETLOG.filter(x=>x.cid===c.id);
  if(!S.pkf||S.pkf.cid!==c.id) S.pkf={cid:c.id,id:nx?nx.id:c.paket,when:'next'};
  const can=canPaket(), f=S.pkf, avail=SERVICES.filter(x=>!ENG.some(e=>e.cid===c.id&&e.sid===x.id));
  return '<div class="dashboard-detail-grid"><section class="card card-pad"><span class="eyebrow">PAKET AKTIF</span><h2>Paket '+esc(p.nama)+' · '+rp(p.harga)+'<small class="note"> / bulan</small></h2><p class="note" style="margin:4px 0 12px">'+esc(p.tag)+'</p>'+chipsOf(c)+'<ul class="pk-list">'+p.fitur.map(x=>'<li>'+esc(x)+'</li>').join('')+'</ul></section>'+
  '<section class="card card-pad"><span class="eyebrow">UBAH PAKET</span><h2 class="sub-title" style="margin-top:6px">Paket untuk '+esc(c.singkat||c.nama)+'</h2>'+
  (nx?'<div class="banner info" style="margin-bottom:12px"><div><strong>Dijadwalkan: '+esc(pk(nx.id).nama)+' mulai '+dfmt(nx.berlaku)+'</strong>Diatur oleh '+esc(nx.by)+'.</div></div>':'')+
  (can?'<div class="col gap-4"><label class="f">Paket baru<select data-change="pksel">'+PAKET.map(x=>'<option value="'+x.id+'"'+(f.id===x.id?' selected':'')+'>'+esc(x.nama)+' · '+rp(x.harga)+' / bulan</option>').join('')+'</select></label><label class="f">Mulai berlaku<select data-change="pkwhen"><option value="next"'+(f.when==='next'?' selected':'')+'>Periode berikutnya (1 Des 2026)</option><option value="now"'+(f.when==='now'?' selected':'')+'>Segera</option></select></label><div class="row"><button class="btn primary" data-a="pkapply">Simpan perubahan paket</button>'+(nx?'<button class="btn" data-a="pknow">Terapkan sekarang</button><button class="btn" data-a="pkcancel">Batalkan jadwal</button>':'')+'</div><p class="note" style="margin:0">Layanan dalam paket baru otomatis ditugaskan kepada tim. Klien menerima notifikasi.</p></div>':'<p class="note">Hanya Managing Director dan tim internal yang dapat mengubah paket klien.</p>')+'</section></div>'+
  '<div class="card"><div class="card-h"><h2>Layanan tambahan di luar paket</h2></div>'+(ad.length?'<div class="tbl-wrap"><table><thead><tr><th>Layanan</th><th>PIC</th><th></th></tr></thead><tbody>'+ad.map(e=>'<tr><td><strong>'+esc(svc(e.sid).nama)+'</strong><div class="note">'+svc(e.sid).siklus+' · '+rp(svc(e.sid).harga)+'</div></td><td>'+esc(UNAME(e.pic))+'</td><td class="r">'+(can?'<button class="btn sm" data-a="addrm" data-id="'+e.id+'">Hentikan</button>':'')+'</td></tr>').join('')+'</tbody></table></div>':'<div class="empty">Tidak ada layanan tambahan.</div>')+(can&&avail.length?'<div class="card-b" style="border-top:1px solid var(--line)"><label class="f" style="max-width:360px">Tambah layanan<select data-change="addadd" data-cid="'+c.id+'"><option value="">Pilih layanan...</option>'+avail.map(x=>'<option value="'+x.id+'">'+esc(x.nama)+' · '+rp(x.harga)+'</option>').join('')+'</select></label></div>':'')+'</div>'+
  '<div class="card card-pad"><span class="eyebrow">RIWAYAT PAKET</span><div class="activity-timeline">'+(log.length?log.map(a=>'<div class="activity-item"><span class="activity-icon">'+ic('list')+'</span><div><b>'+esc(a.text)+'</b><small>'+esc(a.who)+' · '+esc(a.ts)+'</small></div></div>').join(''):'<p class="note">Belum ada perubahan paket.</p>')+'</div></div>';
}
function vPaketCat(){
  const su=isSU();
  return '<div class="tbl-wrap"><table><thead><tr><th>Paket</th><th>Layanan termasuk</th><th class="r">Harga / bulan</th><th class="r">Hemat</th><th>Klien</th></tr></thead><tbody>'+PAKET.map(p=>{ const cs=CLIENTS.filter(c=>c.paket===p.id); return '<tr><td><strong>'+esc(p.nama)+'</strong><div class="note" style="max-width:300px">'+esc(p.tag)+'</div></td><td><div class="chips">'+p.svc.map(id=>'<span class="chip">'+SVSHORT[id]+'</span>').join('')+'</div></td><td class="r">'+(su?'<input type="number" min="0" step="50000" data-input="pkharga" data-id="'+p.id+'" value="'+p.harga+'" style="width:130px;text-align:right">':'<span class="num">'+rp(p.harga)+'</span>')+'</td><td class="r num">'+(paketSaving(p)>0?rp(paketSaving(p)):'–')+'</td><td>'+(cs.length?esc(cs.map(c=>c.singkat||c.nama).join(', ')):'–')+'</td></tr>'; }).join('')+'</tbody></table></div><div class="card-b note">Harga paket berlaku untuk tagihan baru. Invoice yang sudah terbit tidak berubah.</div>';
}
function vKlienTagihan(c){
  if(!(isSU()||isAdminDiv())) return locked('Tagihan hanya untuk Managing Director dan Admin','Hubungi divisi Admin & Penagihan untuk informasi tagihan klien ini.');
  const rows=INVOICES.filter(i=>i.cid===c.id).sort((a,b)=>b.tgl-a.tgl);
  return '<div class="card"><div class="card-h"><h2>Invoice jasa MandaGroup</h2><span class="grow"></span><button class="btn sm primary" data-a="invnew" data-cid="'+c.id+'">'+ic('plus')+'Buat invoice</button></div>'+(rows.length?'<div class="tbl-wrap"><table><thead><tr><th>No. invoice</th><th>Tanggal</th><th>Jatuh tempo</th><th class="r">Total</th><th>Status</th></tr></thead><tbody>'+rows.map(i=>'<tr class="click" data-a="invopen" data-no="'+i.no+'" tabindex="0"><td class="mono">'+i.no+jenisTag(i)+'</td><td class="num">'+dfmt(i.tgl)+'</td><td class="num">'+dfmt(i.due)+'</td><td class="r num"><strong>'+num(invTot(i).total)+'</strong></td><td>'+invStatus(i)+'</td></tr>').join('')+'</tbody></table></div>':'<div class="empty">Belum ada invoice.</div>')+'</div>';
}
function vTim(){
  const tim=USERS.filter(x=>x.role==='tim'), mitra=USERS.filter(x=>x.role==='mitra');
  return head('Tim dan mitra','Penugasan, beban kerja, dan honor mitra pekerjaan.')+
  '<div class="card"><div class="card-h"><h2>Tim internal</h2></div><div class="tbl-wrap"><table><thead><tr><th>Nama</th><th>Divisi</th><th>Jabatan</th><th>Klien dipegang</th><th>Review aktif</th></tr></thead><tbody>'+tim.map(x=>'<tr><td><strong>'+esc(x.nama)+'</strong></td><td>'+DIV[x.div]+'</td><td>'+esc(x.jabatan)+'</td><td class="num">'+new Set(ENG.filter(e=>e.pic===x.id).map(e=>e.cid)).size+'</td><td class="num">'+SUBS.filter(s=>['menunggu','review','qc'].includes(s.status)&&reviewerOf(s)===x.id&&!mitraOf(s)).length+'</td></tr>').join('')+'</tbody></table></div></div>'+
  '<div class="card"><div class="card-h"><h2>Mitra pekerjaan</h2></div><div class="tbl-wrap"><table><thead><tr><th>Nama</th><th>Bidang</th><th>Review aktif</th><th>Selesai</th><th class="r">Honor per review</th><th class="r">Total berjalan</th></tr></thead><tbody>'+mitra.map(x=>{ const a=SUBS.filter(s=>mitraOf(s)===x.id&&['menunggu','review'].includes(s.status)).length, d=SUBS.filter(s=>mitraOf(s)===x.id&&['qc','disetujui','perbaikan'].includes(s.status)&&s.t===10).length; return '<tr><td><strong>'+esc(x.nama)+'</strong></td><td>'+esc(x.jabatan)+'</td><td class="num">'+a+'</td><td class="num">'+d+'</td><td class="r num">'+rp(x.fee)+'</td><td class="r num">'+rp(x.fee*d)+'</td></tr>'; }).join('')+'</tbody></table></div></div><p class="note">Mitra hanya melihat pekerjaan yang ditugaskan kepadanya. Hasil review mitra melewati pengecekan mutu tim Review & Mutu sebelum dirilis ke klien.</p>';
}

/* ---------- invoice ---------- */
function vInvoice(){
  const f=S.invF; const rows=INVOICES.filter(i=>f==='semua'||(f==='terbuka'?(i.status==='terkirim'||i.status==='verifikasi'):i.status===f)).sort((a,b)=>b.tgl-a.tgl||b.no.localeCompare(a.no));
  const out=INVOICES.filter(i=>i.status==='terkirim'||i.status==='verifikasi').reduce((s,i)=>s+invTot(i).total,0), paid=INVOICES.filter(i=>i.status==='lunas').reduce((s,i)=>s+invTot(i).total,0), late=INVOICES.filter(i=>i.status==='terkirim'&&daysTo(i.due)<0).length;
  return head('Invoice klien','Buat, kirim, dan verifikasi pembayaran invoice jasa '+CONFIG.brand+'. Invoice diterbitkan atas nama '+esc(CONFIG.entity)+'.','<button class="btn primary" data-a="invnew">'+ic('plus')+'Buat invoice</button>')+
  '<div class="grid g4">'+tile('Belum dibayar',short(out),'Termasuk menunggu verifikasi')+tile('Lunas',short(paid),'Seluruh periode')+tile('Terlambat',String(late),late?'Kirim pengingat tagihan':'Tidak ada')+tile('Menunggu verifikasi',String(INVOICES.filter(i=>i.status==='verifikasi').length),'Konfirmasi dari klien')+'</div>'+
  '<div class="card"><div class="card-h"><div class="filters">'+[['semua','Semua'],['draft','Draft'],['terbuka','Belum lunas'],['lunas','Lunas']].map(([k,l])=>'<button data-a="invf" data-k="'+k+'" aria-pressed="'+(f===k)+'">'+l+'</button>').join('')+'</div></div><div class="tbl-wrap"><table><thead><tr><th>No. invoice</th><th>Klien</th><th>Tanggal</th><th>Jatuh tempo</th><th class="r">Total</th><th>Status</th></tr></thead><tbody>'+
  rows.map(i=>'<tr class="click" data-a="invopen" data-no="'+i.no+'" tabindex="0"><td class="mono">'+i.no+jenisTag(i)+'</td><td>'+esc(cl(i.cid).nama)+'</td><td class="num">'+dfmt(i.tgl)+'</td><td class="num">'+dfmt(i.due)+'</td><td class="r num"><strong>'+num(invTot(i).total)+'</strong></td><td>'+invStatus(i)+'</td></tr>').join('')+'</tbody></table></div></div>';
}
function vTagihan(){
  const c=cl(me().cid), rows=INVOICES.filter(i=>i.cid===c.id&&i.status!=='draft').sort((a,b)=>b.tgl-a.tgl);
  const out=rows.filter(i=>i.status==='terkirim').reduce((s,i)=>s+invTot(i).total,0);
  return head('Tagihan','Invoice jasa '+CONFIG.brand+' untuk '+esc(c.nama)+'.')+'<div class="grid g4">'+tile('Belum dibayar',out?short(out):'Lunas',out?'Segera lakukan pembayaran':'Tidak ada tagihan')+tile('Total invoice',String(rows.length),'Seluruh periode')+'</div>'+
  '<div class="card"><div class="tbl-wrap"><table><thead><tr><th>No. invoice</th><th>Tanggal</th><th>Jatuh tempo</th><th class="r">Total</th><th>Status</th><th></th></tr></thead><tbody>'+rows.map(i=>'<tr class="click" data-a="invopen" data-no="'+i.no+'" tabindex="0"><td class="mono">'+i.no+jenisTag(i)+'</td><td class="num">'+dfmt(i.tgl)+'</td><td class="num">'+dfmt(i.due)+'</td><td class="r num"><strong>'+num(invTot(i).total)+'</strong></td><td>'+invStatus(i)+'</td><td class="r"><span class="btn sm '+(i.status==='terkirim'?'primary':'')+'">'+(i.status==='terkirim'?'Bayar':'Lihat')+'</span></td></tr>').join('')+'</tbody></table></div></div>';
}
function invoiceHtml(i,plain){ const c=cl(i.cid);
  return invDoc({i,col:'#7A2418',logo:logoMark(54),brand:'<div class="ibrand"><b>Manda<span>Group</span></b><small>'+esc(CONFIG.tagline)+'</small></div><div class="ient">'+esc(CONFIG.entity)+'</div>',lines:[CONFIG.alamat,[CONFIG.email,CONFIG.phone].filter(Boolean).join(' · '),CONFIG.npwp?'NPWP '+CONFIG.npwp:''],to:{n:c.nama,l:['NPWP '+c.npwp]},bank:CONFIG.bank,note:'',status:plain?'':invStatus(i)}); }
const IDOC_CSS=`.idoc{background:#fff;color:#1d1411;padding:28px 30px;border:1px solid #e6dad2;border-radius:12px;display:flex;flex-direction:column;gap:18px;font-size:13px;line-height:1.5}
.idoc .ih{display:flex;gap:14px;align-items:center}.idoc .ih>div:last-child{flex:1}.idoc .ist{align-self:flex-start}
.idoc .ibrand b{font:700 22px var(--f-display,Inter,Arial);letter-spacing:-.01em}.idoc .ibrand b span{color:#D9884F}.idoc .ibrand small{display:block;color:#6b5a52}
.idoc .ient{font-weight:700}.idoc .iln{color:#6b5a52}
.idoc .ittl{text-align:center;font:700 17px var(--f-display,Inter,Arial);letter-spacing:.06em;padding:10px 0;border-top:3px solid var(--ic,#7A2418);border-bottom:1px solid #e6dad2}
.idoc .ipar{display:grid;grid-template-columns:1fr 1fr;gap:20px}.idoc .ipar h3{margin:0 0 4px;font:700 13px var(--f-display,Inter,Arial);color:var(--ic,#7A2418);text-transform:uppercase;letter-spacing:.05em}
.idoc dl.im{display:grid;grid-template-columns:auto 1fr;gap:4px 14px;margin:0}.idoc dl.im dt{font-weight:700}.idoc dl.im dd{margin:0}
.idoc table{width:100%;border-collapse:collapse}.idoc th{background:#f3eeea;text-align:left;padding:9px 10px;font-weight:700;color:#1d1411}.idoc th.r,.idoc td.r{text-align:right;white-space:nowrap}.idoc td{padding:9px 10px;border-bottom:1px solid #efe7e0;color:#1d1411}
.idoc .isum{margin-left:auto;width:min(100%,390px);display:flex;flex-direction:column;gap:5px}.idoc .isum div{display:flex;justify-content:space-between;gap:16px}
.idoc .isum .tot{border-top:2px solid var(--ic,#7A2418);padding-top:8px;margin-top:4px;font:700 15px var(--f-display,Inter,Arial);color:var(--ic,#7A2418)}
.idoc .ibank{line-height:1.7}.idoc .inote{color:#6b5a52;font-size:12px;border-top:1px solid #e6dad2;padding-top:10px}
@media(max-width:640px){.idoc{padding:18px}.idoc .ipar{grid-template-columns:1fr}.idoc .ih{flex-wrap:wrap}}`;
function invDoc(o){ const i=o.i, t=invTot(i), j=t.jenis, nm=x=>x.filter(Boolean);
  const ttl=j==='dp'?'INVOICE DOWN PAYMENT':j==='pelunasan'?'INVOICE PELUNASAN':'INVOICE';
  const meta=[['Invoice No',i.no],['Tanggal',dfmt(i.tgl)],['Jatuh tempo',dfmt(i.due)]]; if(i.ref) meta.push(['Referensi kontrak',i.ref]);
  const sm=[]; if(j!=='penuh') sm.push(['Total kontrak',rp(t.kontrak)]);
  if(j==='dp') sm.push(['Down payment ('+t.pct+'%)',rp(t.dpp)]); else if(j==='pelunasan') sm.push(['Dikurangi down payment','('+rp(t.dpBase)+')'],['Sisa tagihan (DPP)',rp(t.dpp)]); else sm.push(['Subtotal (DPP)',rp(t.dpp)]);
  sm.push([i.ppn?'PPN 11%':'PPN','' +(i.ppn?rp(t.ppn):'Tidak dikenakan')]);
  const last=j==='dp'?'Total tagihan DP':j==='pelunasan'?'Total tagihan pelunasan':'Total tagihan';
  return '<div class="idoc paper" style="--ic:'+o.col+'"><div class="ih">'+o.logo+'<div>'+o.brand+nm(o.lines).map(x=>'<div class="iln">'+esc(x)+'</div>').join('')+'</div></div>'+(o.status?'<div class="ist">'+o.status+'</div>':'')+
   '<div class="ittl">'+ttl+'</div>'+
   '<div class="ipar"><div><h3>Ditagihkan kepada</h3><div class="ient">'+esc(o.to.n)+'</div>'+nm(o.to.l).map(x=>'<div class="iln">'+esc(x)+'</div>').join('')+'</div><dl class="im">'+meta.map(([a,b])=>'<dt>'+a+'</dt><dd>'+esc(b)+'</dd>').join('')+'</dl></div>'+
   '<div><h3 style="margin:0 0 6px;font:700 15px var(--f-display,Inter,Arial)">Rincian layanan</h3><table><thead><tr><th>Deskripsi</th><th class="r">Qty</th><th class="r">Harga</th><th class="r">Total</th></tr></thead><tbody>'+i.lines.map(l=>'<tr><td>'+esc(l.desc)+'</td><td class="r">'+l.qty+'</td><td class="r">'+rp(l.harga)+'</td><td class="r">'+rp(l.qty*l.harga)+'</td></tr>').join('')+'</tbody></table></div>'+
   '<div class="isum">'+sm.map(([a,b])=>'<div><span>'+a+'</span><span>'+b+'</span></div>').join('')+'<div class="tot"><span>'+last+'</span><span>'+rp(t.total)+'</span></div></div>'+
   '<div class="ibank"><strong>Transfer ke:</strong><br>'+esc(o.bank)+'</div>'+
   '<div class="inote">Mohon cantumkan nomor invoice pada berita transfer.'+(j==='pelunasan'?' Tagihan ini telah memperhitungkan down payment yang sudah dibayar.':'')+(o.note?' '+esc(o.note):'')+'</div></div>'; }
function invStandalone(html,title){ return '<!doctype html><html lang="id"><head><meta charset="utf-8"><title>'+esc(title)+'</title><meta name="viewport" content="width=device-width,initial-scale=1"><style>:root{--f-display:Inter,Arial,sans-serif}body{margin:0;padding:24px;background:#f5ede6;font-family:Inter,Arial,sans-serif}.idoc{max-width:820px;margin:0 auto}.lg-mono{display:grid;place-items:center;color:#fff;border-radius:12px;font-weight:700}.lg-img{object-fit:contain;border-radius:8px}'+IDOC_CSS+'@media print{body{background:#fff;padding:0}.idoc{border:0}}</style></head><body>'+html+'</body></html>'; }
function invCopyText(i){ const t=invTot(i); return 'INVOICE '+i.no+'\n'+CONFIG.entity+'\nKepada: '+cl(i.cid).nama+'\nJatuh tempo: '+dfmt(i.due)+'\n\n'+i.lines.map(l=>'- '+l.desc+' ('+l.qty+' x '+rp(l.harga)+')').join('\n')+'\n\nDPP: '+rp(t.dpp)+'\nPPN 11%: '+rp(t.ppn)+'\nTotal: '+rp(t.total)+'\n'+CONFIG.bank; }

/* ---------- pengingat ---------- */
function openSubs(){ return CLIENTS.filter(c=>myClients().some(m=>m.id===c.id)).map(c=>({c,s:SUBS.filter(s=>s.cid===c.id&&(s.status==='belum'||s.status==='perbaikan')).sort((a,b)=>b.t-a.t)[0]})).filter(x=>x.s); }
function nextReminder(c,s){
  if(!c.kalender) return 'Nonaktif';
  const due=dueOf(s); if(daysTo(due)<0) return RULES.find(r=>r.k==='late').on?'Otomatis tiap 3 hari sampai terkirim':'Tidak ada aturan aktif';
  const c2=RULES.filter(r=>r.on&&r.off!==null).map(r=>new Date(due.getTime()+r.off*864e5)).filter(d=>d>=TODAY).sort((a,b)=>a-b)[0];
  return c2?dfmt(c2):'Tidak ada aturan aktif';
}
function vPengingat(){
  const list=openSubs(), su=isSU();
  return head('Pengingat dan kalender kepatuhan','Pengingat otomatis agar klien mengirim laporan tepat waktu dan menanggapi temuan.','<button class="btn primary" data-a="remall">'+ic('bell')+'Kirim ke semua yang terlambat</button>')+
  '<div class="grid g2"><div class="card"><div class="card-h"><h2>Status pengingat per klien</h2></div><div class="tbl-wrap"><table><thead><tr><th>Klien</th><th>Periode</th><th>Batas</th><th>Pengingat berikutnya</th><th></th></tr></thead><tbody>'+
   (list.length?list.map(({c,s})=>'<tr><td><strong>'+esc(c.nama)+'</strong></td><td>'+periodLabel(s.t)+(s.status==='perbaikan'?' '+pill('bad','Perbaikan'):'')+'</td><td class="num">'+dfmt(dueOf(s))+'<div>'+(daysTo(dueOf(s))<0?pill('bad','Lewat '+(-daysTo(dueOf(s)))+' hari'):pill('mute',daysTo(dueOf(s))+' hari lagi'))+'</div></td><td>'+nextReminder(c,s)+'</td><td class="r"><button class="btn sm" data-a="remsend" data-cid="'+c.id+'">Kirim sekarang</button></td></tr>').join(''):'<tr><td colspan="5" class="empty">Semua klien sudah mengirim laporan.</td></tr>')+'</tbody></table></div></div>'+
  '<div class="card"><div class="card-h"><h2>Aturan pengingat</h2></div><div class="list">'+RULES.map(r=>'<label class="item"><input type="checkbox" data-change="ruletoggle" data-k="'+r.k+'"'+(r.on?' checked':'')+(su?'':' disabled')+'><div class="grow"><strong>'+r.label+'</strong><small>'+r.ch+'</small></div></label>').join('')+'</div>'+(su?'':'<div class="card-b note">Hanya superuser yang dapat mengubah aturan.</div>')+'</div></div>'+
  '<div class="card"><div class="card-h"><h2>Log pengingat terkirim</h2></div><div class="list">'+(REMLOG.length?REMLOG.filter(r=>myClients().some(c=>c.id===r.cid)).map(r=>'<div class="item"><span class="mono" style="min-width:100px">'+r.ts+'</span><div class="grow"><strong>'+esc(cl(r.cid).nama)+'</strong><small>'+esc(r.text)+'</small></div><span class="note">'+r.ch+'</span></div>').join(''):'<div class="empty">Belum ada pengingat.</div>')+'</div></div>';
}

/* ---------- dokumen, pesan, notifikasi, pengaturan ---------- */
function vDokumen(){
  const u=me(); const clients=myClients(); if(!S.docCid||!clients.some(c=>c.id===S.docCid)) S.docCid=clients[0]&&clients[0].id;
  if(!clients.length) return head('Dokumen','Kirim dokumen ke tim '+CONFIG.brand+' dan unduh dokumen hasil kerja kami.')+emptyBox('Belum ada klien. Tambahkan klien terlebih dahulu di menu Klien untuk mengelola dokumen.');
  const cid=isKlien()?u.cid:S.docCid, docs=DOCS.filter(d=>d.cid===cid).sort((a,b)=>b.tgl-a.tgl);
  return head('Dokumen','Kirim dokumen ke tim '+CONFIG.brand+' dan unduh dokumen hasil kerja kami.')+
  '<div class="card"><div class="card-h">'+(isKlien()?'':'<label class="f" style="flex-direction:row;align-items:center;gap:8px">Klien<select data-change="doccid">'+clients.map(c=>'<option value="'+c.id+'"'+(c.id===cid?' selected':'')+'>'+esc(c.nama)+'</option>').join('')+'</select></label>')+'<span class="grow"></span><label class="f" style="flex-direction:row;align-items:center;gap:8px">Kategori<select id="upCat"><option>Dokumen sumber</option><option>Hasil review</option><option>Lainnya</option></select></label><label class="btn primary" style="cursor:pointer">'+ic('upload')+'Unggah dokumen<input type="file" data-change="upload" multiple hidden></label></div>'+
  '<div class="tbl-wrap"><table><thead><tr><th>Nama berkas</th><th>Kategori</th><th>Tanggal</th><th>Diunggah oleh</th><th class="r">Ukuran</th></tr></thead><tbody>'+docs.map(d=>'<tr><td>'+esc(d.nama)+'</td><td>'+pill(d.kat==='Hasil review'?'info':'mute',esc(d.kat))+'</td><td class="num">'+dfmt(d.tgl)+'</td><td>'+esc(UNAME(d.by))+'</td><td class="r num">'+esc(d.size)+'</td></tr>').join('')+'</tbody></table></div></div><p class="note">Pada versi demo, hanya nama berkas yang tercatat. Penyimpanan terenkripsi di server tersedia setelah integrasi dengan backend.</p>';
}
function vPesan(){
  const u=me(), clients=myClients(); if(!S.msgCid||!clients.some(c=>c.id===S.msgCid)) S.msgCid=clients[0]&&clients[0].id;
  if(!clients.length) return head('Pesan','Percakapan dengan klien.')+emptyBox('Belum ada klien. Tambahkan klien terlebih dahulu di menu Klien untuk memulai percakapan.');
  const cid=isKlien()?u.cid:S.msgCid, th=THREADS[cid]=THREADS[cid]||[], mine=isKlien()?'klien':'tim';
  th.forEach(m=>{ if(m.side!==mine) m.read=true; });
  const list=isKlien()?'':'<div class="card"><div class="thr">'+clients.map(c=>'<button data-a="msgsel" data-cid="'+c.id+'" aria-current="'+(c.id===cid)+'"><strong>'+esc(c.nama)+'</strong><span class="note">'+esc(((THREADS[c.id]||[]).slice(-1)[0]||{text:'Belum ada pesan'}).text.slice(0,48))+'</span></button>').join('')+'</div></div>';
  return head('Pesan',isKlien()?'Percakapan dengan tim '+CONFIG.brand+'.':'Percakapan dengan klien.')+'<div class="'+(isKlien()?'':'split')+'">'+list+'<div class="card"><div class="card-h"><h2>'+esc(cl(cid).nama)+'</h2></div><div class="chat" id="chat">'+(th.length?th.map(m=>'<div class="msg '+(m.side===mine?'me':'')+'"><small>'+esc(m.name)+' · '+m.ts+'</small>'+esc(m.text)+'</div>').join(''):'<div class="empty">Belum ada pesan.</div>')+'</div><div class="card-b" style="border-top:1px solid var(--line)"><label class="f">Pesan baru<textarea id="draft" data-input="draft" placeholder="Tulis pesan...">'+esc(S.draft)+'</textarea></label><div class="row" style="margin-top:10px"><button class="btn primary" data-a="send">Kirim pesan</button></div></div></div></div>';
}
function vNotif(){
  const n=myNotifs();
  return head('Notifikasi','Pembaruan status laporan, pengingat, dan tagihan.','<button class="btn" data-a="notifall">Tandai semua dibaca</button>')+
  '<div class="card">'+(n.length?'<div class="list">'+n.map(x=>'<button class="item itemlink '+(x.read?'':'stripe info')+'" data-a="notifopen" data-id="'+x.id+'"><div class="grow"><strong'+(x.read?' style="font-weight:500"':'')+'>'+esc(x.text)+'</strong><small>'+x.ts+'</small></div>'+(x.read?'':pill('info','Baru'))+'</button>').join('')+'</div>':'<div class="empty">Tidak ada notifikasi.</div>')+'</div>';
}
function vPengaturan(){
  const caps=[['Melihat semua klien dan layanan','✓','–','–','–'],['Menugaskan tim dan mitra','✓','–','–','–'],['Mereview laporan klien','✓','✓*','✓','–'],['Merilis hasil ke klien (QC)','✓','✓*','–','–'],['Membuat dan memverifikasi invoice','✓','✓**','–','–'],['Mengatur aturan pengingat','✓','–','–','–'],['Mengirim laporan dan dokumen','–','–','–','✓'],['Menanggapi temuan dan chat','–','✓','–','✓'],['Melihat invoice dan honor sendiri','✓','✓**','honor','✓']];
  return head('Pengaturan','Peran, hak akses, dan kebijakan layanan.')+
  '<div class="card"><div class="card-h"><h2>Hak akses per peran</h2></div><div class="tbl-wrap"><table class="matrix"><thead><tr><th>Kemampuan</th><th>Superuser</th><th>Tim</th><th>Mitra</th><th>Klien</th></tr></thead><tbody>'+caps.map(r=>'<tr>'+r.map(x=>'<td>'+x+'</td>').join('')+'</tr>').join('')+'</tbody></table></div><div class="card-b note">* Divisi Review & Mutu, atau reviewer yang ditugaskan. ** Divisi Admin & Penagihan. Tim hanya melihat klien yang ditugaskan kepadanya; mitra hanya melihat pekerjaan yang ditugaskan.</div></div>'+
  '<div class="card"><div class="card-h"><h2>Profil penerbit invoice</h2></div><div class="card-b"><div class="formgrid"><label class="f" style="grid-column:1/-1">Nama perusahaan<input type="text" value="'+esc(CONFIG.entity)+'" disabled></label><label class="f" style="grid-column:1/-1">Alamat<input type="text" data-input="cfg" data-k="alamat" value="'+esc(CONFIG.alamat)+'" placeholder="Belum diisi"></label><label class="f">Email<input type="text" data-input="cfg" data-k="email" value="'+esc(CONFIG.email)+'"></label><label class="f">Telepon<input type="text" data-input="cfg" data-k="phone" value="'+esc(CONFIG.phone)+'" placeholder="Belum diisi"></label><label class="f">NPWP<input type="text" data-input="cfg" data-k="npwp" value="'+esc(CONFIG.npwp)+'" placeholder="Belum diisi"></label><label class="f" style="grid-column:1/-1">Rekening pembayaran<input type="text" data-input="cfg" data-k="bank" value="'+esc(CONFIG.bank)+'" placeholder="Bank, nomor rekening, atas nama"></label></div><p class="note" style="margin:10px 0 0">Data ini tampil pada kop setiap invoice MandaGroup. Alamat dan telepon sengaja dikosongkan dan tidak ditampilkan pada invoice sampai diisi. Domain: '+esc(CONFIG.domain)+'.</p></div></div>'+
  '<div class="card"><div class="card-h"><h2>Kebijakan layanan</h2></div><div class="card-b"><dl class="kv"><dt>Batas kirim laporan</dt><dd>Tanggal '+CONFIG.deadlineDay+' bulan berikutnya</dd><dt>SLA review</dt><dd>'+SERVICES[0].sla+' hari kerja sejak laporan diterima</dd><dt>Alur hasil mitra</dt><dd>Mitra mereview, tim Review & Mutu merilis ke klien</dd><dt>PPN invoice</dt><dd>11%, dapat dimatikan per invoice</dd><dt>Penerbit invoice</dt><dd>'+esc(CONFIG.entity)+'</dd></dl></div></div>';
}

/* ---------- modal ---------- */
function modalHtml(){
  const m=S.modal; if(!m) return '';
  const close='<button class="btn sm" data-a="close" aria-label="Tutup">'+ic('x')+'</button>', wrap=(t,b)=>'<div class="overlay" data-a="overlay"><div class="modal" role="dialog" aria-modal="true" aria-label="'+t+'"><div class="card-h"><h2>'+t+'</h2><span class="grow"></span>'+close+'</div><div class="card-b" style="display:flex;flex-direction:column;gap:14px">'+b+'</div></div></div>';
  if(m.t==='inv'){
    const i=INVOICES.find(x=>x.no===m.no), adm=isAdminDiv(); let a='';
    if(isKlien()&&i.status==='terkirim') a='<button class="btn primary" data-a="invpay" data-no="'+i.no+'">Saya sudah transfer</button>';
    if(adm&&i.status==='draft') a='<button class="btn primary" data-a="invsend" data-no="'+i.no+'">Kirim ke klien</button>';
    if(adm&&i.status==='verifikasi') a='<button class="btn primary" data-a="invverify" data-no="'+i.no+'">Verifikasi dan tandai lunas</button>';
    if(adm&&i.status==='terkirim') a='<button class="btn" data-a="invremind" data-no="'+i.no+'">Kirim pengingat tagihan</button><button class="btn" data-a="invverify" data-no="'+i.no+'">Tandai lunas</button>';
    return wrap('Invoice '+i.no,invoiceHtml(i)+'<div class="row">'+a+'<button class="btn" data-a="cinvprint">Cetak atau simpan PDF</button><button class="btn" data-a="invdl" data-no="'+i.no+'">Unduh invoice (HTML)</button><button class="btn" data-a="invcopy" data-no="'+i.no+'">'+ic('copy')+'Salin ringkasan</button></div>');
  }
  if(m.t==='invnew'){
    const d=S.invd, t=invTot(invD(S.invd));
    return wrap('Buat invoice',
     '<div class="row"><label class="f" style="flex:1">Klien<select data-change="invcid">'+CLIENTS.map(c=>'<option value="'+c.id+'"'+(d.cid===c.id?' selected':'')+'>'+esc(c.nama)+'</option>').join('')+'</select></label><label class="f">Jatuh tempo (hari)<input type="number" min="1" id="invDays" data-input="invdays" value="'+d.days+'" style="width:110px"></label></div>'+
     jenisFields(d,'inv',dpList(INVOICES,d.cid))+'<div>'+d.lines.map((l,n)=>'<div class="row" style="margin-bottom:8px"><label class="f" style="flex:1 1 220px">Uraian<input type="text" data-input="invline" data-i="'+n+'" data-f="desc" value="'+esc(l.desc)+'"></label><label class="f">Jml<input type="number" min="1" data-input="invline" data-i="'+n+'" data-f="qty" value="'+l.qty+'" style="width:70px"></label><label class="f">Harga (Rp)<input type="number" min="0" data-input="invline" data-i="'+n+'" data-f="harga" value="'+l.harga+'" style="width:140px"></label><button class="btn sm" data-a="invrm" data-i="'+n+'" aria-label="Hapus baris">'+ic('x')+'</button></div>').join('')+'</div>'+
     '<div class="row"><label class="f">Tambah dari katalog<select data-change="invadd"><option value="">Pilih layanan...</option>'+SERVICES.map(s=>'<option value="'+s.id+'">'+esc(s.nama)+' · '+rp(s.harga)+'</option>').join('')+'</select></label><button class="btn sm" data-a="invcustom">'+ic('plus')+'Baris manual</button><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" data-change="invppn"'+(d.ppn?' checked':'')+'> Kenakan PPN 11%</label></div>'+
     '<div class="card-b card" id="invTot">'+totHtml(t)+'</div><div class="row"><button class="btn" data-a="invsave" data-s="draft">Simpan draft</button><button class="btn primary" data-a="invsave" data-s="terkirim">Terbitkan ke klien</button></div>');
  }
  if(m.t==='cinv'){
    const i=CINV.find(x=>x.no===m.no); if(!i) return '';
    const b=[]; if(i.status==='draft') b.push('<button class="btn primary" data-a="cinvsend" data-no="'+i.no+'">'+ic('send')+'Tandai terkirim</button><button class="btn" data-a="cinvedit" data-no="'+i.no+'">Ubah</button><button class="btn" data-a="cinvdel" data-no="'+i.no+'">Hapus draft</button>');
    if(i.status==='terkirim') b.push('<button class="btn primary" data-a="cinvpaid" data-no="'+i.no+'">'+ic('check')+'Tandai sudah dibayar</button>');
    return wrap('Invoice '+i.no,cinvHtml(i)+'<div class="row">'+b.join('')+'<button class="btn" data-a="cinvprint">Cetak atau simpan PDF</button><button class="btn" data-a="cinvdl" data-no="'+i.no+'">Unduh invoice (HTML)</button><button class="btn" data-a="cinvcopy" data-no="'+i.no+'">'+ic('copy')+'Salin ringkasan</button></div>');
  }
  if(m.t==='cinvnew'){
    const d=S.cin, c=cl(me().cid), t=invTot(invD(d));
    return wrap(d.no?'Ubah invoice '+d.no:'Buat invoice usaha',
     '<div class="kopmini">'+logoBox(c,34)+'<span><b>'+esc(c.nama)+'</b><small>Kop invoice memakai logo dan identitas dari Profil usaha</small></span></div>'+
     '<div class="formgrid"><label class="f" style="grid-column:1/-1">Nama pelanggan<input type="text" data-input="cin" data-k="cust" value="'+esc(d.cust)+'" placeholder="Nama perusahaan atau perorangan"></label><label class="f" style="grid-column:1/-1">Alamat pelanggan<input type="text" data-input="cin" data-k="addr" value="'+esc(d.addr)+'"></label><label class="f">Email pelanggan<input type="text" data-input="cin" data-k="email" value="'+esc(d.email)+'"></label><label class="f">Jatuh tempo (hari)<input type="number" min="1" data-input="cindays" value="'+d.days+'"></label></div>'+
     jenisFields(d,'cin',dpList(CINV,c.id))+'<div>'+d.lines.map((l,n)=>'<div class="row" style="margin-bottom:8px"><label class="f" style="flex:1 1 220px">Uraian<input type="text" data-input="cinline" data-i="'+n+'" data-f="desc" value="'+esc(l.desc)+'"></label><label class="f">Jml<input type="number" min="1" data-input="cinline" data-i="'+n+'" data-f="qty" value="'+l.qty+'" style="width:70px"></label><label class="f">Harga (Rp)<input type="number" min="0" data-input="cinline" data-i="'+n+'" data-f="harga" value="'+l.harga+'" style="width:140px"></label><button class="btn sm" data-a="cinrm" data-i="'+n+'" aria-label="Hapus baris">'+ic('x')+'</button></div>').join('')+'</div>'+
     '<div class="row"><button class="btn sm" data-a="cinadd">'+ic('plus')+'Tambah baris</button><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" data-change="cinppn"'+(d.ppn?' checked':'')+'> Kenakan PPN 11% (bila usaha Anda PKP)</label></div>'+
     '<label class="f">Catatan untuk pelanggan<input type="text" data-input="cin" data-k="note" value="'+esc(d.note)+'"></label>'+
     '<div class="card-b card" id="cinTot">'+totHtml(t)+'</div><div class="row"><button class="btn" data-a="cinsave" data-s="draft">Simpan draft</button><button class="btn primary" data-a="cinsave" data-s="terkirim">Simpan dan tandai terkirim</button></div>');
  }
  if(m.t==='kxexp'){ const d=S.kxd; return wrap('Catat pengeluaran','<div class="formgrid"><label class="f">Tanggal<input type="date" data-input="kxd" data-k="tgl" value="'+d.tgl+'"></label><label class="f">Kategori<select data-input="kxd" data-k="kat">'+Object.entries(KXCAT).map(([k,v])=>'<option value="'+k+'"'+(d.kat===k?' selected':'')+'>'+v[0]+'</option>').join('')+'</select></label><label class="f" style="grid-column:1/-1">Keterangan<input type="text" data-input="kxd" data-k="ket" value="'+esc(d.ket)+'" placeholder="Contoh: Sewa coworking November"></label><label class="f">Jumlah (Rp)<input type="number" min="0" data-input="kxd" data-k="jml" value="'+d.jml+'"></label></div><div class="row"><button class="btn primary" data-a="kxsave">Simpan pengeluaran</button><button class="btn" data-a="close">Batal</button></div>'); }
  if(m.t==='reopen') return wrap('Buka kembali '+periodLabel(m.tm),'<p class="note" style="margin:0">Alasan akan tercatat di jejak audit dan klien diberi tahu.</p><label class="f">Alasan pembukaan<textarea data-input="reopen" placeholder="Contoh: koreksi jurnal penyesuaian yang terlewat">'+esc(S.reopen)+'</textarea></label><div class="row"><button class="btn primary" data-a="reopensave">Buka kembali periode</button><button class="btn" data-a="close">Batal</button></div>');
  if(m.t==='newclient') return wrap('Tambah klien','<div class="row"><label class="f" style="flex:1 1 100%">Nama perusahaan<input type="text" id="nc_nama" value="'+esc(S.nc.nama)+'"></label><label class="f">Bentuk<select id="nc_bentuk"><option>PT</option><option>CV</option><option>Perorangan</option><option>Yayasan</option></select></label><label class="f" style="flex:1">Bidang usaha<input type="text" id="nc_usaha" value="'+esc(S.nc.usaha)+'"></label></div><p class="note" style="margin:0">Setelah dibuat, tugaskan layanan dan tim pada halaman klien.</p><div class="row"><button class="btn primary" data-a="ncsave">Simpan klien</button></div>');
  return '';
}
const totHtml=t=>{ const r=(a,b)=>'<dt>'+a+'</dt><dd class="num">'+b+'</dd>'; let h=''; if(t.jenis!=='penuh') h+=r('Total kontrak',rp(t.kontrak)); if(t.jenis==='dp') h+=r('Down payment ('+t.pct+'%)',rp(t.dpp)); else if(t.jenis==='pelunasan') h+=r('Dikurangi down payment','('+rp(t.dpBase)+')')+r('Sisa tagihan (DPP)',rp(t.dpp)); else h+=r('DPP',rp(t.dpp));
  return '<dl class="kv" style="grid-template-columns:1fr auto;justify-items:end;margin:0">'+h+r('PPN 11%',rp(t.ppn))+'<dt><strong>Total tagihan</strong></dt><dd class="num"><strong>'+rp(t.total)+'</strong></dd></dl>'; };
function jenisFields(d,p,list){ const j=d.jenis||'penuh', sel=(v,l)=>'<option value="'+v+'"'+(j===v?' selected':'')+'>'+l+'</option>';
  return '<div class="row" style="align-items:flex-end"><label class="f">Jenis invoice<select data-change="'+p+'jenis">'+sel('penuh','Tagihan penuh')+sel('dp','Down payment (DP)')+sel('pelunasan','Pelunasan')+'</select></label>'+
   (j!=='penuh'?'<label class="f" style="flex:1 1 200px">Referensi kontrak<input type="text" data-input="'+p+'x" data-k="ref" value="'+esc(d.ref||'')+'" placeholder="Contoh: KONTRAK/2026/014"></label>':'')+
   (j==='dp'||(j==='pelunasan'&&!d.dpRef)?'<label class="f">DP (%)<input type="number" min="1" max="99" data-input="'+p+'x" data-k="dpPct" value="'+(d.dpPct||30)+'" style="width:90px"></label>':'')+
   (j==='pelunasan'?'<label class="f" style="flex:1 1 220px">Invoice DP yang sudah dibayar<select data-change="'+p+'dpref"><option value="">Hitung dari persentase DP</option>'+list.map(x=>'<option value="'+esc(x.no)+'"'+(d.dpRef===x.no?' selected':'')+'>'+esc(x.no)+' · '+rp(x.amt)+'</option>').join('')+'</select></label>':'')+'</div>'+
   (j!=='penuh'?'<p class="note" style="margin:0">Isi baris dengan nilai kontrak penuh. '+(j==='dp'?'Sistem menagih '+(d.dpPct||30)+'% dari total kontrak.':'Sistem mengurangi down payment yang sudah dibayar dari total kontrak.')+'</p>':''); }
const dpList=(arr,cid)=>arr.filter(x=>x.cid===cid&&x.jenis==='dp'&&x.status!=='draft').map(x=>({no:x.no,amt:invTot(x).dpp}));

/* ---------- render dan aksi ---------- */
function render(){
  try{ $('app').classList.toggle('is-land',!S.entered); }catch(e){}
  if(!S.entered){ $('app').innerHTML=landing(); return; }
  const old=$('content'), sc=old?old.scrollTop:0;
  $('app').innerHTML=shell();
  const n=$('content'); if(n) n.scrollTop=sc;
  const ch=$('chat'); if(ch) ch.scrollTop=ch.scrollHeight;
  if(S.view==='kirim'&&!S.sub&&isKlien()&&S.form&&$('chk')){ const r=chkHtml(); $('chk').innerHTML=r.html; $('chkSum').innerHTML=r.sum; }
}
let tt; function toast(t){ const e=$('toast'); e.textContent=t; e.hidden=false; clearTimeout(tt); tt=setTimeout(()=>e.hidden=true,2800); }
function copyText(t,ok){ try{ navigator.clipboard.writeText(t).then(()=>toast(ok),()=>toast('Salin gagal. Pilih teks lalu salin manual.')); }catch(e){ toast('Salin gagal. Pilih teks lalu salin manual.'); } }
const hist=(s,text,internal)=>s.hist.push({ts:stamp(),who:me().nama,text,internal:!!internal});
function goView(v,sub){ S.view=v; S.sub=sub||null; S.modal=null; if(v!=='klien') S.cd=null; render(); const n=$('content'); if(n) n.scrollTop=0; }
function release(s,v){ s.status=v==='sesuai'?'disetujui':'perbaikan'; s.pend=null; s.doneAt=TODAY; if(s.status==='perbaikan') s.fixDue=new Date(TODAY.getTime()+3*864e5); hist(s,'Hasil review dirilis: '+(v==='sesuai'?'sesuai':'perlu perbaikan ('+s.findings.filter(f=>f.st==='terbuka').length+' temuan terbuka)')); notify('c:'+s.cid,v==='sesuai'?'Laporan '+periodLabel(s.t)+' dinyatakan sesuai.':'Hasil review '+periodLabel(s.t)+' tersedia: laporan perlu perbaikan.','hasil',s.id); }
function submitForm(){
  const f=S.form, s=SUBS.find(x=>x.id===f.sid), r=chkHtml(), miss=DOCREQ.filter(d=>!f.files[d.k]);
  if(miss.length) return toast('Lengkapi dokumen wajib terlebih dahulu: '+miss.map(x=>x.nama).join(', '));
  if(r.checks.some(x=>x.lvl==='err'&&x.id!=='doc')&&!f.catatan.trim()) return toast('Ada ketidaksesuaian angka. Tuliskan penjelasan pada kolom catatan.');
  const wasFix=s.status==='perbaikan';
  s.fig=Object.fromEntries(FIG.map(([k])=>[k,+f.f[k]||0])); s.files={...f.files}; s.catatan=f.catatan; s.sentAt=TODAY; s.status='menunggu'; s.draft=false; if(wasFix) s.rev++;
  hist(s,wasFix?'Laporan dikirim ulang (revisi ke-'+s.rev+')':'Laporan '+periodLabel(s.t)+' dikirim');
  notify('u:'+reviewerOf(s),cl(s.cid).nama+' mengirim '+(wasFix?'ulang ':'')+'laporan '+periodLabel(s.t)+' untuk direview.','antrean',s.id);
  if(mitraOf(s)) notify('u:'+mitraOf(s),'Laporan '+periodLabel(s.t)+' '+cl(s.cid).nama+' siap direview.','antrean',s.id);
  S.form=null; toast('Laporan terkirim. Anda akan menerima notifikasi saat hasil review tersedia.'); goView('hasil',s.id);
}
function act(a,d){
  const u=me();
  switch(a){
    case 'lgfill': { const x=USERS.find(y=>y.id===d.id); if(!x) return; S.lg={email:demoEmail(x),pw:DEMO_PW,err:'',show:S.lg.show}; return render(); }
    case 'lgshow': S.lg.show=!S.lg.show; return render();
    case 'lgforgot': return toast('Hubungi '+CONFIG.email+(CONFIG.phone?' atau '+CONFIG.phone:'')+' untuk mengatur ulang akses.');
    case 'lgsubmit': { const em=String(S.lg.email).trim().toLowerCase(); if(!em||!S.lg.pw){ S.lg.err='Email dan kata sandi wajib diisi.'; return render(); } const x=USERS.find(y=>demoEmail(y)===em); if(!x||S.lg.pw!==DEMO_PW){ S.lg.err='Email atau kata sandi tidak sesuai.'; return render(); } S.lg={email:'',pw:'',err:'',show:false}; return act('enter',{id:x.id}); }
    case 'enter': S.uid=d.id; S.entered=true; S.view='dashboard'; S.sub=null; S.cd=null; S.modal=null; S.form=null; S.msgCid=null; S.docCid=null; return render();
    case 'theme': S.theme=isDark()?'light':'dark'; try{ localStorage.setItem(THEME_KEY,S.theme); }catch(e){} applyTheme(); return render();
    case 'refresh': toast('Data diperbarui'); return render();
    case 'logout': S.entered=false; S.modal=null; return render();
    case 'go': return goView(d.v,d.sub);
    case 'sub': return goView(S.view,d.id);
    case 'back': S.sub=null; return render();
    case 'close': S.modal=null; return render();
    case 'qf': S.qF=d.k; return render();
    case 'ltab': S.ltab=d.k; return render();
    case 'invf': S.invF=d.k; return render();
    case 'cd': S.cd=d.id; S.cdt='ringkas'; return render();
    case 'cdback': S.cd=null; return render();
    case 'reopen': { const s=SUBS.find(x=>x.id===d.id); openForm(s); return goView('kirim'); }
    case 'startrev': { const s=SUBS.find(x=>x.id===d.id); s.status='review'; if(u.role==='tim') s.reviewer=u.id; hist(s,'Review dimulai'); notify('c:'+s.cid,'Laporan '+periodLabel(s.t)+' sedang direview oleh tim kami.','hasil',s.id); return render(); }
    case 'addfind': { const s=SUBS.find(x=>x.id===d.id); if(!S.fnd.teks.trim()) return toast('Tulis uraian temuan terlebih dahulu'); s.findings.push({id:'F'+(s.findings.length+1),area:S.fnd.area,sev:S.fnd.sev,teks:S.fnd.teks.trim(),st:'terbuka',resp:''}); S.fnd.teks=''; hist(s,'Temuan ditambahkan ('+S.fnd.area+')',true); return render(); }
    case 'closef': { const s=SUBS.find(x=>x.id===d.id), f=s.findings.find(x=>x.id===d.fid); f.st='ditutup'; hist(s,'Temuan '+f.id+' ditutup',true); return render(); }
    case 'reopenf': { const s=SUBS.find(x=>x.id===d.id), f=s.findings.find(x=>x.id===d.fid); f.st='terbuka'; hist(s,'Temuan '+f.id+' dikembalikan untuk ditanggapi ulang'); return render(); }
    case 'verdict': { const s=SUBS.find(x=>x.id===d.id), v=d.v; if(v==='sesuai'&&s.findings.some(f=>f.st!=='ditutup')) return toast('Masih ada temuan yang belum ditutup'); if(v==='perbaikan'&&!s.findings.some(f=>f.st==='terbuka')) return toast('Tambahkan minimal satu temuan terbuka'); if(u.role==='mitra'){ s.status='qc'; s.pend=v; hist(s,'Review selesai oleh mitra: usulan '+(v==='sesuai'?'sesuai':'minta perbaikan')); notify('u:'+reviewerOf(s),'Hasil review mitra '+cl(s.cid).nama+' '+periodLabel(s.t)+' menunggu rilis mutu.','antrean',s.id); toast('Hasil dikirim ke tim untuk pengecekan mutu'); } else { release(s,v); toast('Hasil dirilis ke klien'); } return render(); }
    case 'release': { const s=SUBS.find(x=>x.id===d.id); release(s,s.pend||'perbaikan'); toast('Hasil dirilis ke klien'); return render(); }
    case 'retmitra': { const s=SUBS.find(x=>x.id===d.id); s.status='review'; s.pend=null; hist(s,'Dikembalikan ke mitra untuk diperbaiki',true); if(mitraOf(s)) notify('u:'+mitraOf(s),'Hasil review '+cl(s.cid).nama+' dikembalikan oleh tim mutu.','antrean',s.id); return render(); }
    case 'addnote': { const s=SUBS.find(x=>x.id===d.id); if(!S.note.trim()) return toast('Tulis catatan terlebih dahulu'); hist(s,S.note.trim(),true); S.note=''; return render(); }
    case 'answer': { const s=SUBS.find(x=>x.id===d.id), f=s.findings.find(x=>x.id===d.fid), t=(S.reply[s.id+f.id]||'').trim(); if(!t) return toast('Tulis tanggapan terlebih dahulu'); f.resp=t; f.st='dijawab'; hist(s,'Tanggapan temuan '+f.id+' dikirim'); const left=s.findings.filter(x=>x.st==='terbuka').length; toast(left?'Tanggapan tersimpan. Masih ada '+left+' temuan terbuka.':'Semua temuan sudah ditanggapi. Kirim ulang laporan agar direview.'); if(!left){ openForm(s); S.view='kirim'; S.sub=null; } return render(); }
    case 'formsave': { const s=SUBS.find(x=>x.id===S.form.sid), f=S.form; s.fig=Object.fromEntries(FIG.map(([k])=>[k,+f.f[k]||0])); s.files={...f.files}; s.catatan=f.catatan; if(s.status==='belum') s.draft=true; toast('Draft tersimpan'); return render(); }
    case 'formsubmit': return submitForm();
    case 'akb': S.akb=d.k; return render();
    case 'akgo': S.akb=d.k; return goView('akun');
    case 'akopen': S.akc=d.cid; S.akb='laporan'; return goView('akun');
    case 'prhor': S.prhor=+d.k; return render();
    case 'prmet': S.prmet=d.k; return render();
    case 'prsc': S.prsc=d.k; return render();
    case 'akprint': try{ document.body.classList.add('rpt'); window.print(); }catch(e){ toast('Gunakan Ctrl+P untuk mencetak atau simpan sebagai PDF.'); } try{ document.body.classList.remove('rpt'); }catch(e){} return;
    case 'akcsv': { const cid=akCid(), t=S.akt; downloadText('mandagroup-'+cid+'-'+periodLabel(t).replace(/\s/g,'-').toLowerCase()+'.csv',csvOf(cid,t)); audit(cid,'Laporan','Mengunduh laporan '+periodLabel(t)+' (CSV)'); return; }
    case 'fillfig': { const s=SUBS.find(x=>x.id===S.form.sid), cid=u.cid; if(!TRX.some(x=>x.cid===cid&&x.t===s.t)) return toast('Belum ada transaksi pada '+periodLabel(s.t)+'. Catat transaksi di menu Laporan dan analisis.'); S.form.f=figLedger(cid,s.t); audit(cid,'Laporan','Mengisi angka laporan '+periodLabel(s.t)+' dari buku besar'); toast('Angka terisi otomatis dari buku besar'); return render(); }
    case 'trxadd': { const cid=akCid(), f=S.trxf, t=+f.t, jml=Math.round(+f.jml), dd=Math.min(28,Math.max(1,Math.round(+f.d)||1));
      if(!canInputAk(cid)) return toast('Anda tidak memiliki izin mencatat transaksi');
      if(isClosed(cid,t)) return toast('Periode '+periodLabel(t)+' sudah ditutup');
      if(!String(f.ket).trim()) return toast('Isi keterangan transaksi');
      if(!(jml>0)) return toast('Jumlah harus lebih dari nol');
      if(!f.deb||!f.kre||f.deb===f.kre) return toast('Pilih akun debit dan kredit yang berbeda');
      if(!accOf(cid,f.deb)||!accOf(cid,f.kre)) return toast('Akun tidak dikenal');
      addTrx({cid,t,d:dd,ket:String(f.ket).trim(),deb:f.deb,kre:f.kre,jml,by:u.nama,ts:stamp()});
      audit(cid,'Transaksi','Mencatat "'+String(f.ket).trim()+'" sebesar '+rp(jml)+' ('+f.deb+' / '+f.kre+')');
      f.ket=''; f.jml=''; f.d=''; toast('Transaksi dicatat'); return render(); }
    case 'trxdel': { const x=TRX.find(y=>y.id===d.id); if(!x||!canInputAk(x.cid)) return; if(isClosed(x.cid,x.t)) return toast('Periode sudah ditutup'); TRX.splice(TRX.indexOf(x),1); audit(x.cid,'Transaksi','Menghapus "'+x.ket+'" sebesar '+rp(x.jml)); toast('Transaksi dihapus'); return render(); }
    case 'accadd': { const cid=akCid(), f=S.accf; if(!canStaffAk(cid)) return toast('Anda tidak memiliki izin menambah akun'); const k=String(f.k).trim(); if(!/^\d{4}$/.test(k)) return toast('Kode akun harus 4 digit angka'); if(accOf(cid,k)) return toast('Kode akun sudah dipakai'); if(!String(f.n).trim()) return toast('Isi nama akun'); const a={k,n:String(f.n).trim(),t:f.t}; if(f.t==='aset'||f.t==='liab') a.g='lancar'; if(f.t==='beban') a.g='opx'; (ACCX[cid]=ACCX[cid]||[]).push(a); audit(cid,'Akun','Menambah akun '+k+' '+a.n); S.accf={k:'',n:'',t:'beban'}; toast('Akun ditambahkan'); return render(); }
    case 'bkcatat': { const x=BANKX.find(y=>y.id===d.id); if(!x||!canInputAk(x.cid)) return; if(isClosed(x.cid,x.t)) return toast('Periode sudah ditutup'); const out=x.jml<0; addTrx({cid:x.cid,t:x.t,d:Math.min(28,x.d),ket:x.ket+' (rekonsiliasi bank)',deb:out?'6300':'1100',kre:out?'1100':'4100',jml:Math.abs(x.jml),by:u.nama,ts:stamp()}); BANKX.splice(BANKX.indexOf(x),1); audit(x.cid,'Rekonsiliasi','Mencatat ke buku: '+x.ket+' ('+rp(x.jml)+')'); toast('Item dicatat ke buku'); return render(); }
    case 'bkinv': { const x=BANKX.find(y=>y.id===d.id); if(!x||!canReviewAk(x.cid)||isClosed(x.cid,x.t)) return; x.st='diselidiki'; audit(x.cid,'Rekonsiliasi','Menandai untuk diselidiki: '+x.ket); notify('c:'+x.cid,'Mutasi bank "'+x.ket+'" sedang diselidiki tim. Mohon siapkan keterangannya.','akun'); toast('Ditandai untuk diselidiki'); return render(); }
    case 'afadd': { const cid=akCid(), f=S.afs; if(!canReviewAk(cid)) return toast('Anda tidak memiliki izin memberi temuan'); if(!String(f.teks).trim()) return toast('Tulis uraian temuan'); FSEQ++; ACCFIND.unshift({id:'AF'+FSEQ,cid,k:f.k,t:S.akt,sev:f.sev,teks:String(f.teks).trim(),st:'terbuka',resp:'',by:u.nama,ts:stamp()}); audit(cid,'Review','Temuan pada akun '+f.k+' (tingkat '+f.sev+')'); notify('c:'+cid,'Ada temuan baru pada akun '+f.k+' '+(accOf(cid,f.k)||{n:''}).n+'. Mohon ditanggapi.','akun'); f.teks=''; toast('Temuan ditambahkan dan dikirim ke klien'); return render(); }
    case 'afresp': { const f=ACCFIND.find(x=>x.id===d.id), tx=String(S.afr[d.id]||'').trim(); if(!f||!isKlien()||f.cid!==u.cid) return; if(!tx) return toast('Tulis tanggapan Anda terlebih dahulu'); f.resp=tx; f.st='ditanggapi'; delete S.afr[d.id]; audit(f.cid,'Review','Menanggapi temuan akun '+f.k); notifyPic(f.cid,cl(f.cid).nama+' menanggapi temuan akun '+f.k+'.'); toast('Tanggapan terkirim'); return render(); }
    case 'afclose': { const f=ACCFIND.find(x=>x.id===d.id); if(!f||!canReviewAk(f.cid)) return; f.st='ditutup'; audit(f.cid,'Review','Menutup temuan akun '+f.k); notify('c:'+f.cid,'Temuan akun '+f.k+' telah diterima dan ditutup.','akun'); toast('Temuan ditutup'); return render(); }
    case 'noteadd': { const cid=akCid(), f=S.notef; if(!canReviewAk(cid)) return toast('Anda tidak memiliki izin menambah catatan'); if(!String(f.judul).trim()||!String(f.teks).trim()) return toast('Isi judul dan isi catatan'); CSEQ++; NOTES.unshift({id:'C'+CSEQ,cid,judul:String(f.judul).trim(),teks:String(f.teks).trim(),st:'baru',internal:!!f.internal,by:u.nama,ts:stamp()}); audit(cid,'Catatan','Menambah catatan "'+String(f.judul).trim()+'"'+(f.internal?' (internal)':'')); if(!f.internal) notify('c:'+cid,'Catatan baru dari konsultan: '+String(f.judul).trim(),'akun'); S.notef={judul:'',teks:'',internal:false}; toast('Catatan disimpan'); return render(); }
    case 'notest': { const n=NOTES.find(x=>x.id===d.id); if(!n) return; const ok=isKlien()?(n.cid===u.cid&&!n.internal):canReviewAk(n.cid); if(!ok||!['dalam proses','selesai'].includes(d.s)) return; n.st=d.s; audit(n.cid,'Catatan','Status catatan "'+n.judul+'" menjadi '+d.s); if(isKlien()) notifyPic(n.cid,cl(n.cid).nama+' memperbarui status catatan "'+n.judul+'".'); else if(!n.internal) notify('c:'+n.cid,'Status catatan "'+n.judul+'" diperbarui: '+d.s+'.','akun'); return render(); }
    case 'tgadd': { const cid=akCid(), f=S.tgf, nv=Math.round(+f.nilai); if(!(isKlien()||canStaffAk(cid))) return; if(!(nv>0)) return toast('Isi target lebih dari nol'); GSEQ++; TARGETS.push({id:'G'+GSEQ,cid,metrik:f.metrik,nilai:nv}); audit(cid,'Target','Menambah target '+METRIK[f.metrik]+' sebesar '+rp(nv)); S.tgf.nilai=''; toast('Target ditambahkan'); return render(); }
    case 'tgdel': { const g=TARGETS.find(x=>x.id===d.id); if(!g||!(isKlien()?g.cid===u.cid:canStaffAk(g.cid))) return; TARGETS.splice(TARGETS.indexOf(g),1); audit(g.cid,'Target','Menghapus target '+METRIK[g.metrik]); return render(); }
    case 'sesadd': { const cid=akCid(), f=S.sesf; if(!(isKlien()||canStaffAk(cid))) return; const dt=new Date(f.tgl+'T00:00:00'); if(!f.tgl||isNaN(dt)) return toast('Pilih tanggal konsultasi'); if(dt<TODAY) return toast('Pilih tanggal yang belum lewat'); if(!String(f.topik).trim()) return toast('Isi topik konsultasi'); SSEQ++; SESI.push({id:'S'+SSEQ,cid,tgl:dt,topik:String(f.topik).trim(),st:isKlien()?'diajukan':'dijadwalkan',pic:isKlien()?'':u.nama}); audit(cid,'Target',(isKlien()?'Mengajukan':'Menjadwalkan')+' sesi konsultasi "'+String(f.topik).trim()+'" pada '+dfmt(dt)); if(isKlien()) notifyPic(cid,cl(cid).nama+' mengajukan sesi konsultasi pada '+dfmt(dt)+'.'); else notify('c:'+cid,'Sesi konsultasi dijadwalkan pada '+dfmt(dt)+': '+String(f.topik).trim()+'.','akun'); S.sesf={tgl:'',topik:''}; toast(isKlien()?'Pengajuan sesi terkirim':'Sesi dijadwalkan'); return render(); }
    case 'sesok': { const x=SESI.find(y=>y.id===d.id); if(!x||!canReviewAk(x.cid)) return; x.st='dijadwalkan'; x.pic=u.nama; audit(x.cid,'Target','Mengonfirmasi sesi konsultasi '+dfmt(x.tgl)); notify('c:'+x.cid,'Sesi konsultasi '+dfmt(x.tgl)+' dikonfirmasi oleh '+u.nama+'.','akun'); toast('Sesi dikonfirmasi'); return render(); }
    case 'sesbatal': { const x=SESI.find(y=>y.id===d.id); if(!x||!(isKlien()?x.cid===u.cid:canReviewAk(x.cid))) return; x.st='batal'; audit(x.cid,'Target','Membatalkan sesi konsultasi '+dfmt(x.tgl)); toast('Sesi dibatalkan'); return render(); }
    case 'akclose': { const cid=akCid(), t=+d.t; if(!(isSU()||(isTim()&&me().div==='review'))) return toast('Hanya Managing Director atau tim Review yang dapat menutup buku'); const sb=SUBS.find(x=>x.cid===cid&&x.t===t); if(!(sb&&sb.status==='disetujui')) return toast('Periode dapat ditutup setelah laporan disetujui'); CLOSED[cid+'|'+t]={by:u.nama,ts:stamp()}; audit(cid,'Periode','Menutup buku '+periodLabel(t)); notify('c:'+cid,'Buku '+periodLabel(t)+' telah ditutup. Angka periode ini terkunci.','akun'); toast('Periode '+periodLabel(t)+' ditutup'); return render(); }
    case 'akreopen': { if(!isSU()) return toast('Hanya Managing Director yang dapat membuka kembali periode'); S.reopen=''; S.modal={t:'reopen',cid:akCid(),tm:+d.t}; return render(); }
    case 'reopensave': { const m=S.modal; if(!isSU()||!m||m.t!=='reopen') return; const r=String(S.reopen||'').trim(); if(r.length<5) return toast('Tulis alasan pembukaan kembali (minimal 5 karakter)'); delete CLOSED[m.cid+'|'+m.tm]; audit(m.cid,'Periode','Membuka kembali '+periodLabel(m.tm)+'. Alasan: '+r); notify('c:'+m.cid,'Periode '+periodLabel(m.tm)+' dibuka kembali oleh Managing Director.','akun'); S.modal=null; S.reopen=''; toast('Periode dibuka kembali'); return render(); }
    case 'cdt': S.cdt=d.k; return render();
    case 'pkapply': { const c=cl(S.cd), f=S.pkf; if(!canPaket()) return toast('Anda tidak memiliki izin mengubah paket'); if(f.id===c.paket&&!c.paketNext) return toast('Paket yang dipilih sama dengan paket aktif'); const from=pk(c.paket).nama, to=pk(f.id).nama;
      if(f.id===c.paket){ c.paketNext=null; toast('Jadwal perubahan paket dibatalkan'); return render(); }
      if(f.when==='now'){ applyPaket(c,f.id); PAKETLOG.unshift({ts:stamp(),cid:c.id,who:u.nama,text:'Paket diubah dari '+from+' ke '+to+' (berlaku segera)'}); notify('c:'+c.id,'Paket layanan Anda diubah dari '+from+' ke '+to+' dan berlaku mulai sekarang.','layanan'); toast('Paket diubah ke '+to); }
      else { c.paketNext={id:f.id,berlaku:T(11,1),by:u.nama}; PAKETLOG.unshift({ts:stamp(),cid:c.id,who:u.nama,text:'Perubahan paket ke '+to+' dijadwalkan berlaku 1 Des 2026'}); notify('c:'+c.id,'Paket layanan Anda akan berubah ke '+to+' mulai 1 Des 2026.','layanan'); toast('Perubahan paket dijadwalkan'); }
      S.pkf=null; return render(); }
    case 'pknow': { const c=cl(S.cd); if(!canPaket()||!c.paketNext) return; const to=pk(c.paketNext.id).nama, from=pk(c.paket).nama; applyPaket(c,c.paketNext?c.paketNext.id:c.paket); PAKETLOG.unshift({ts:stamp(),cid:c.id,who:u.nama,text:'Paket diubah dari '+from+' ke '+to+' (diterapkan lebih awal)'}); notify('c:'+c.id,'Paket layanan Anda sekarang '+to+'.','layanan'); S.pkf=null; toast('Paket diterapkan'); return render(); }
    case 'pkcancel': { const c=cl(S.cd); if(!canPaket()) return; c.paketNext=null; PAKETLOG.unshift({ts:stamp(),cid:c.id,who:u.nama,text:'Jadwal perubahan paket dibatalkan'}); S.pkf=null; toast('Jadwal dibatalkan'); return render(); }
    case 'addrm': { const e=ENG.find(x=>x.id===d.id); if(!canPaket()||!e) return; ENG=ENG.filter(x=>x.id!==d.id); PAKETLOG.unshift({ts:stamp(),cid:e.cid,who:u.nama,text:'Layanan tambahan dihentikan: '+svc(e.sid).nama}); toast('Layanan tambahan dihentikan'); return render(); }
    case 'reqpaket': { const c=cl(u.cid), to=pk(d.id).nama; notify('u:md','Klien '+c.nama+' meminta perubahan paket ke '+to+'.','klien'); toast('Permintaan perubahan ke paket '+to+' dikirim. Tim kami akan menghubungi Anda.'); return; }
    case 'cinvf': S.cinvF=d.k; return render();
    case 'cinvopen': S.modal={t:'cinv',no:d.no}; return render();
    case 'cinvnew': S.cin={no:null,cust:'',addr:'',email:'',days:14,ppn:false,note:'',jenis:'penuh',ref:'',dpPct:30,dpRef:'',lines:[{desc:'',qty:1,harga:0}]}; S.modal={t:'cinvnew'}; return render();
    case 'cinvedit': { const i=CINV.find(x=>x.no===d.no); S.cin={no:i.no,cust:i.cust,addr:i.addr,email:i.email,days:Math.max(1,Math.round((i.due-i.tgl)/864e5)),ppn:i.ppn,note:i.note,jenis:i.jenis||'penuh',ref:i.ref||'',dpPct:i.dpPct||30,dpRef:i.dpRef||'',lines:i.lines.map(l=>({...l}))}; S.modal={t:'cinvnew'}; return render(); }
    case 'cinadd': S.cin.lines.push({desc:'',qty:1,harga:0}); return render();
    case 'cinrm': S.cin.lines.splice(+d.i,1); if(!S.cin.lines.length) S.cin.lines.push({desc:'',qty:1,harga:0}); return render();
    case 'cinsave': { const x=S.cin, c=cl(u.cid); if(!x.cust.trim()) return toast('Isi nama pelanggan'); if(x.lines.some(l=>!l.desc.trim())) return toast('Lengkapi uraian pada semua baris'); const lines=x.lines.map(l=>({desc:l.desc.trim(),qty:+l.qty||1,harga:+l.harga||0}));
      if(x.no){ const i=CINV.find(v=>v.no===x.no); Object.assign(i,{cust:x.cust.trim(),addr:x.addr,email:x.email,note:x.note,ppn:x.ppn,jenis:x.jenis,ref:(x.ref||'').trim(),dpPct:+x.dpPct||30,dpRef:x.dpRef||'',lines,due:new Date(i.tgl.getTime()+x.days*864e5),status:d.s}); }
      else mkCinv({cid:c.id,cust:x.cust.trim(),addr:x.addr,email:x.email,note:x.note,ppn:x.ppn,jenis:x.jenis,ref:(x.ref||'').trim(),dpPct:+x.dpPct||30,dpRef:x.dpRef||'',lines,tgl:TODAY,due:new Date(TODAY.getTime()+x.days*864e5),status:d.s});
      S.modal=null; S.cin=null; toast(d.s==='terkirim'?'Invoice disimpan dan ditandai terkirim. Cetak atau salin untuk dikirim ke pelanggan.':'Draft invoice disimpan'); return render(); }
    case 'cinvsend': { CINV.find(x=>x.no===d.no).status='terkirim'; toast('Invoice ditandai terkirim'); return render(); }
    case 'cinvpaid': { const ci=CINV.find(x=>x.no===d.no); ci.status='lunas'; ci.tglBayar=TODAY; S.modal=null; toast('Invoice ditandai sudah dibayar'); return render(); }
    case 'cinvdel': { const k=CINV.findIndex(x=>x.no===d.no); if(k>=0) CINV.splice(k,1); S.modal=null; toast('Draft dihapus'); return render(); }
    case 'cinvcopy': return copyText(cinvText(CINV.find(x=>x.no===d.no)),'Ringkasan invoice disalin');
    case 'cinvprint': try{ window.print(); }catch(e){ toast('Gunakan Ctrl+P untuk mencetak.'); } return;
    case 'profsave': toast('Profil usaha disimpan. Logo dan identitas dipakai pada semua invoice.'); return;
    case 'profcolor': cl(u.cid).brand.warna=d.c; return render();
    case 'logorm': cl(u.cid).brand.logo=null; return render();
    case 'fkc': S.fkc=d.k; S.ftab='lapor'; return render();
    case 'fkp': S.fkp=+d.k; return render();
    case 'fst': S.fst=d.k; return render();
    case 'ftab': S.ftab=d.k; return render();
    case 'finreq': notify('u:md',u.nama+' meminta akses data keuangan '+cl(d.cid).nama+'.','keuangan'); toast('Permintaan akses dikirim ke Managing Director'); return;
    case 'invopen': S.modal={t:'inv',no:d.no}; return render();
    case 'invdl': { const i=INVOICES.find(x=>x.no===d.no); return downloadText('Invoice-'+i.no.replace(/[^\w-]/g,'_')+'.html',invStandalone(invoiceHtml(i,true),'Invoice '+i.no)); }
    case 'cinvdl': { const i=CINV.find(x=>x.no===d.no); return downloadText('Invoice-'+i.no.replace(/[^\w-]/g,'_')+'.html',invStandalone(cinvHtml(i,true),'Invoice '+i.no)); }
    case 'kt': S.kt=d.k; return render();
    case 'kxnew': S.kxd={tgl:'2026-11-13',kat:'sewa',ket:'',jml:''}; S.modal={t:'kxexp'}; return render();
    case 'kxsave': { const x=S.kxd, p=String(x.tgl).split('-').map(Number), jml=+x.jml; if(!x.ket.trim()||!(jml>0)) return toast('Isi keterangan dan jumlah pengeluaran'); const tg=new Date(p[0],p[1]-1,p[2]); if(isNaN(tg)||tg.getFullYear()!==2026||!AKM.includes(tg.getMonth()+1)||tg>TODAY) return toast('Tanggal harus antara Juli 2026 dan hari ini'); kxAdd(tg,x.kat,x.ket.trim(),jml,true); S.modal=null; toast('Pengeluaran dicatat'); return render(); }
    case 'kxdel': { const n=KXEXP.findIndex(x=>x.id===d.id&&x.user); if(n>=0){ KXEXP.splice(n,1); toast('Pengeluaran dihapus'); } return render(); }
    case 'kxcsv': return downloadText('keuangan-manda.csv',kxCsv());
    case 'kxprint': try{ document.body.classList.add('rpt'); window.print(); }catch(e){ toast('Gunakan Ctrl+P untuk mencetak atau simpan sebagai PDF.'); } try{ document.body.classList.remove('rpt'); }catch(e){} return;
    case 'invnew': { if(!CLIENTS.length) return toast('Tambahkan klien terlebih dahulu sebelum membuat invoice'); const cid0=d.cid||CLIENTS[0].id; S.invd={cid:cid0,days:14,ppn:true,auto:true,jenis:'penuh',ref:'',dpPct:30,dpRef:'',lines:paketLines(cl(cid0))}; }  S.modal={t:'invnew'}; return render();
    case 'invadd': return;
    case 'invcustom': S.invd.lines.push({desc:'',qty:1,harga:0}); return render();
    case 'invrm': S.invd.lines.splice(+d.i,1); if(!S.invd.lines.length) S.invd.lines.push({desc:'',qty:1,harga:0}); return render();
    case 'invsave': { const x=S.invd; if(x.lines.some(l=>!l.desc.trim())) return toast('Lengkapi uraian pada semua baris'); const tgl=TODAY; const i=mkInv({cid:x.cid,tgl,due:new Date(TODAY.getTime()+x.days*864e5),lines:x.lines.map(l=>({desc:l.desc.trim(),qty:+l.qty||1,harga:+l.harga||0})),ppn:x.ppn,jenis:x.jenis,ref:(x.ref||'').trim(),dpPct:+x.dpPct||30,dpRef:x.dpRef||'',status:d.s,by:u.id}); if(d.s==='terkirim') notify('c:'+x.cid,'Invoice baru '+i.no+' ('+rp(invTot(i).total)+') jatuh tempo '+dfmt(i.due)+'.','tagihan'); S.modal=null; toast(d.s==='terkirim'?'Invoice '+i.no+' dikirim ke klien':'Draft invoice tersimpan'); return render(); }
    case 'invsend': { const i=INVOICES.find(x=>x.no===d.no); i.status='terkirim'; notify('c:'+i.cid,'Invoice baru '+i.no+' ('+rp(invTot(i).total)+') jatuh tempo '+dfmt(i.due)+'.','tagihan'); S.modal=null; toast('Invoice dikirim ke klien'); return render(); }
    case 'invpay': { const i=INVOICES.find(x=>x.no===d.no); i.status='verifikasi'; notify('u:md','Klien '+cl(i.cid).nama+' mengonfirmasi pembayaran '+i.no+'.','invoice'); S.modal=null; toast('Konfirmasi terkirim. Tim kami akan memverifikasi pembayaran Anda.'); return render(); }
    case 'invverify': { const i=INVOICES.find(x=>x.no===d.no); i.status='lunas'; i.tglBayar=TODAY; notify('c:'+i.cid,'Pembayaran invoice '+i.no+' telah diverifikasi. Terima kasih.','tagihan'); S.modal=null; toast('Invoice ditandai lunas'); return render(); }
    case 'invremind': { const i=INVOICES.find(x=>x.no===d.no); notify('c:'+i.cid,'Pengingat tagihan: invoice '+i.no+' ('+rp(invTot(i).total)+') jatuh tempo '+dfmt(i.due)+'.','tagihan'); S.modal=null; toast('Pengingat tagihan terkirim'); return render(); }
    case 'invcopy': return copyText(invCopyText(INVOICES.find(x=>x.no===d.no)),'Ringkasan invoice disalin');
    case 'remsend': { const c=cl(d.cid), s=SUBS.filter(x=>x.cid===c.id&&(x.status==='belum'||x.status==='perbaikan')).sort((a,b)=>b.t-a.t)[0]; if(!s) return toast(c.nama+' tidak memiliki laporan terbuka'); notify('c:'+c.id,'Pengingat: '+(s.status==='perbaikan'?'tanggapi temuan dan kirim ulang laporan ':'kirim laporan keuangan ')+periodLabel(s.t)+' paling lambat '+dfmt(dueOf(s))+'.',s.status==='perbaikan'?'hasil':'kirim',s.status==='perbaikan'?s.id:null); REMLOG.push({ts:stamp(),cid:c.id,text:'Pengingat manual oleh '+u.nama+': '+periodLabel(s.t),ch:'Portal, email'}); toast('Pengingat dikirim ke '+c.nama); return render(); }
    case 'remall': { let n=0; openSubs().forEach(({c,s})=>{ if(s.status==='belum'&&daysTo(dueOf(s))<0&&c.kalender){ notify('c:'+c.id,'Pengingat keterlambatan: laporan '+periodLabel(s.t)+' belum kami terima (batas '+dfmt(dueOf(s))+').','kirim'); REMLOG.push({ts:stamp(),cid:c.id,text:'Pengingat keterlambatan manual: '+periodLabel(s.t),ch:'Portal, email, WhatsApp'}); n++; } }); toast(n?n+' pengingat terkirim':'Tidak ada klien yang terlambat'); return render(); }
    case 'openmsg': S.msgCid=d.cid; return goView('pesan');
    case 'msgsel': S.msgCid=d.cid; return render();
    case 'send': { const t=S.draft.trim(); if(!t) return toast('Tulis pesan terlebih dahulu'); const cid=isKlien()?u.cid:S.msgCid; (THREADS[cid]=THREADS[cid]||[]).push({side:isKlien()?'klien':'tim',name:u.nama,text:t,ts:stamp(),read:false}); if(isKlien()){ const e=engOf(cid,'S01'); notify('u:'+(e?e.pic:'md'),'Pesan baru dari '+cl(cid).nama+'.','pesan'); } else notify('c:'+cid,'Pesan baru dari '+u.nama+'.','pesan'); S.draft=''; return render(); }
    case 'notifopen': { const n=NOTIFS.find(x=>x.id===d.id); n.read=true; return goView(n.view||'dashboard',n.sub); }
    case 'notifall': myNotifs().forEach(n=>n.read=true); return render();
    case 'reqsvc': { const s=svc(d.id), c=cl(u.cid); notify('u:md','Pengajuan layanan: '+c.nama+' meminta '+s.nama+'.','klien'); toast('Pengajuan terkirim. Tim kami akan menghubungi Anda.'); return render(); }
    case 'newclient': S.modal={t:'newclient'}; return render();
    case 'ncsave': { const nama=$('nc_nama').value.trim(); if(!nama) return toast('Nama perusahaan wajib diisi'); const id='c'+(CLIENTS.length+1); const pre=(nama.replace(/^(PT|CV|UD)\.?\s+/i,'').split(/\s+/).map(w=>w[0]||'').join('').replace(/[^A-Za-z]/g,'').toUpperCase().slice(0,3))||'INV'; const nc={id,nama,bentuk:$('nc_bentuk').value,usaha:$('nc_usaha').value||'-',npwp:'-',kalender:true,sejak:'Nov 2026',paket:'review',paketNext:null,finAccess:[],brand:{logo:null,seq:0,catatan:'Terima kasih atas kepercayaan Anda.',alamat:'',email:'',telp:'',rek:'',pre,warna:'#7A2418'}}; CLIENTS.push(nc); applyPaket(nc,'review'); TAXM[id]='badan'; PKP[id]=false; S.modal=null; S.cd=id; toast('Klien ditambahkan'); return render(); }
    case 'overlay': if(d.__t) { S.modal=null; render(); } return;
  }
}
document.addEventListener('click',e=>{
  const el=e.target.closest('[data-a]'); if(!el) return;
  const a=el.dataset.a, d={...el.dataset};
  if(a==='overlay'){ if(e.target===el){ S.modal=null; render(); } return; }
  act(a,d);
});
document.addEventListener('keydown',e=>{ if(e.key==='Enter'&&e.target&&e.target.dataset&&e.target.dataset.input==='lg'){ e.preventDefault(); act('lgsubmit',{}); return; }  if((e.key==='Enter'||e.key===' ')&&e.target.matches&&e.target.matches('tr.click')){ e.preventDefault(); e.target.click(); } if(e.key==='Escape'&&S.modal){ S.modal=null; render(); } });
function onChange(el){
  const k=el.dataset.change; if(!k) return; const v=el.value;
  switch(k){
    case 'persona': S.uid=v; S.view='dashboard'; S.sub=null; S.cd=null; S.modal=null; S.form=null; S.msgCid=null; S.docCid=null; S.qF='semua'; return render();
    case 'formper': { const s=SUBS.find(x=>x.id===v); openForm(s); return render(); }
    case 'formfile': { const f=el.files&&el.files[0]; if(!f) return; S.form.files[el.dataset.k]=f.name; return render(); }
    case 'engpic': { const e=ENG.find(x=>x.id===el.dataset.id); e.pic=v; toast('PIC diperbarui: '+UNAME(v)); return render(); }
    case 'engmitra': { const e=ENG.find(x=>x.id===el.dataset.id); e.mitra=v||null; if(v) notify('u:'+v,'Anda ditugaskan untuk '+cl(e.cid).nama+' ('+svc(e.sid).nama+').','antrean'); toast(v?'Mitra ditugaskan: '+UNAME(v):'Mitra dilepas dari penugasan'); return render(); }
    case 'submitra': { const s=SUBS.find(x=>x.id===el.dataset.id); s.mitra=v||null; hist(s,v?'Ditugaskan ke mitra '+UNAME(v):'Dikerjakan tim internal',true); if(v) notify('u:'+v,'Anda ditugaskan mereview '+cl(s.cid).nama+' '+periodLabel(s.t)+'.','antrean',s.id); return render(); }
    case 'akc': S.akc=v; return render();
    case 'akt': S.akt=+v; S.trxf.t=+v; return render();
    case 'akf': { const f=el.dataset.f; if(!AKF.includes(f)) return; const kk=el.dataset.k; S[f][kk]=el.type==='checkbox'?el.checked:v; if(f==='trxf'&&kk==='jenis'){ const j=JENIS.find(x=>x[0]===v); if(j&&j[2]){ S.trxf.deb=j[2]; S.trxf.kre=j[3]; } } return render(); }
    case 'bud': { const cid=akCid(); if(!canStaffAk(cid)) return; const k=el.dataset.k, nv=Math.max(0,Math.round(+v)||0), b=budOf(cid); if(b[k]===nv) return; audit(cid,'Anggaran','Mengubah anggaran '+(accOf(cid,k)||{n:k}).n+' dari '+rp(b[k]||0)+' menjadi '+rp(nv)); b[k]=nv; return render(); }
    case 'asm': { const cid=akCid(); if(!canStaffAk(cid)) return; const k=el.dataset.k, a=pa(cid), nv=+v; if(!isFinite(nv)||v==='') return; audit(cid,'Proyeksi','Mengubah asumsi '+({g:'pertumbuhan pendapatan',vp:'beban variabel',fx:'beban tetap'}[k]||k)+' dari '+a[k]+' menjadi '+nv); a[k]=nv; return render(); }
    case 'taxm': { const cid=akCid(); if(!canStaffAk(cid)) return; TAXM[cid]=v; audit(cid,'Pajak','Mengubah skema pajak penghasilan menjadi '+(v==='final'?'PPh Final UMKM 0,5%':'PPh Badan 22%')); return render(); }
    case 'taxpkp': { const cid=akCid(); if(!canStaffAk(cid)) return; PKP[cid]=!!el.checked; audit(cid,'Pajak','Status PKP diubah menjadi '+(PKP[cid]?'PKP':'bukan PKP')); return render(); }
    case 'afk': S.afk=v; return render();
    case 'afq': S.afq=v; return render();
    case 'pksel': S.pkf.id=v; return;
    case 'pkwhen': S.pkf.when=v; return;
    case 'addadd': { if(!v||!canPaket()) return; const c=cl(el.dataset.cid); ENG.push({id:'E'+(++ESEQ),cid:c.id,sid:v,pic:defPic(c.id,v),mitra:null}); PAKETLOG.unshift({ts:stamp(),cid:c.id,who:me().nama,text:'Layanan tambahan ditambahkan: '+svc(v).nama}); notify('c:'+c.id,'Layanan tambahan aktif untuk Anda: '+svc(v).nama+'.','layanan'); toast('Layanan tambahan ditambahkan'); return render(); }
    case 'cinppn': S.cin.ppn=el.checked; return render();
    case 'logoup': { const f=el.files&&el.files[0]; if(!f) return; if(!/^image\//.test(f.type)) return toast('Pilih berkas gambar'); if(f.size>600*1024) return toast('Ukuran logo maksimal 600 KB'); const r=new FileReader(); r.onload=()=>{ cl(me().cid).brand.logo=r.result; toast('Logo diperbarui'); render(); }; r.readAsDataURL(f); return; }
    case 'finacc': { const c=cl(el.dataset.cid), x=USERS.find(v=>v.id===el.dataset.uid); const on=el.checked; c.finAccess=c.finAccess.filter(id=>id!==x.id); if(on) c.finAccess.push(x.id); FINLOG.unshift({ts:stamp(),who:me().nama,text:(on?'Memberi':'Mencabut')+' akses data keuangan '+c.nama+(on?' kepada ':' dari ')+x.nama}); notify('c:'+c.id,'Managing Director '+(on?'memberi':'mencabut')+' akses data keuangan Anda '+(on?'kepada ':'dari ')+x.nama+'.','usaha'); notify('u:'+x.id,'Akses data keuangan '+c.nama+(on?' diberikan kepada Anda.':' dicabut.'),on?'keuangan':'dashboard'); toast('Hak akses diperbarui'); return render(); }
    case 'caltoggle': cl(el.dataset.cid).kalender=el.checked; return render();
    case 'ruletoggle': RULES.find(r=>r.k===el.dataset.k).on=el.checked; return render();
    case 'doccid': S.docCid=v; return render();
    case 'fnd_area': S.fnd.area=v; return;
    case 'fnd_sev': S.fnd.sev=v; return;
    case 'invjenis': S.invd.jenis=v; S.invd.dpRef=''; return render();
    case 'cinjenis': S.cin.jenis=v; S.cin.dpRef=''; return render();
    case 'invdpref': S.invd.dpRef=v; return render();
    case 'cindpref': S.cin.dpRef=v; return render();
    case 'ktp': S.ktp=+v; return render();
    case 'invcid': S.invd.cid=v; S.invd.dpRef=''; if(S.invd.auto) S.invd.lines=paketLines(cl(v)); return render();
    case 'invppn': S.invd.ppn=el.checked; return render();
    case 'invadd': { if(!v) return; const s=svc(v); S.invd.lines.push({desc:s.nama,qty:1,harga:s.harga}); return render(); }
    case 'upload': { const cid=isKlien()?me().cid:S.docCid, kat=($('upCat')||{}).value||'Dokumen sumber'; [...el.files].forEach(x=>DOCS.push({cid,nama:x.name,kat,tgl:TODAY,by:me().id,size:Math.max(1,Math.round(x.size/1024))+' KB'})); if(isKlien()) notify('u:'+((engOf(cid,'S02')||engOf(cid,'S01')).pic),cl(cid).nama+' mengunggah '+el.files.length+' dokumen.','dokumen'); toast(el.files.length+' berkas ditambahkan'); return render(); }
  }
}
document.addEventListener('change',e=>onChange(e.target));
document.addEventListener('input',e=>{
  const el=e.target, k=el.dataset.input; if(!k) return;
  if(k==='fld'){ S.form.f[el.dataset.k]=el.value; const r=chkHtml(); $('chk').innerHTML=r.html; $('chkSum').innerHTML=r.sum; }
  else if(k==='catatan'){ S.form.catatan=el.value; const r=chkHtml(); $('chk').innerHTML=r.html; $('chkSum').innerHTML=r.sum; }
  else if(k==='draft') S.draft=el.value;
  else if(k==='note') S.note=el.value;
  else if(k==='fnd_teks') S.fnd.teks=el.value;
  else if(k==='reply') S.reply[S.sub+el.dataset.fid]=el.value;
  else if(k==='akf'&&AKF.includes(el.dataset.f)){ S[el.dataset.f][el.dataset.k]=el.value; }
  else if(k==='lg'){ S.lg[el.dataset.k]=el.value; S.lg.err=''; }
  else if(k==='afr'){ S.afr[el.dataset.id]=el.value; }
  else if(k==='reopen'){ S.reopen=el.value; }
  else if(k==='pkharga'){ pk(el.dataset.id).harga=+el.value||0; }
  else if(k==='prof'){ const c=cl(me().cid); c.brand[el.dataset.k]=el.value; const kp=$('kopPrev'); if(kp) kp.innerHTML=kopHtml(c); }
  else if(k==='cin') S.cin[el.dataset.k]=el.value;
  else if(k==='cindays') S.cin.days=+el.value||14;
  else if(k==='cinline'){ const l=S.cin.lines[+el.dataset.i]; l[el.dataset.f]=el.dataset.f==='desc'?el.value:(+el.value||0); const t=$('cinTot'); if(t) t.innerHTML=totHtml(invTot(invD(S.cin))); }
  else if(k==='invx'||k==='cinx'){ const D=k==='invx'?S.invd:S.cin; D[el.dataset.k]=el.dataset.k==='dpPct'?Math.min(99,Math.max(1,+el.value||30)):el.value; const t=$(k==='invx'?'invTot':'cinTot'); if(t) t.innerHTML=totHtml(invTot(invD(D))); }
  else if(k==='kxd'){ S.kxd[el.dataset.k]=el.value; }
  else if(k==='cfg'){ CONFIG[el.dataset.k]=el.value; }
  else if(k==='invdays') S.invd.days=+el.value||14;
  else if(k==='invline'){ S.invd.auto=false; const l=S.invd.lines[+el.dataset.i]; l[el.dataset.f]=el.dataset.f==='desc'?el.value:(+el.value||0); const t=$('invTot'); if(t) t.innerHTML=totHtml(invTot(invD(S.invd))); }
});
S.lg={email:'',pw:'',err:'',show:false};
S.theme=themePref(); applyTheme();
render();
