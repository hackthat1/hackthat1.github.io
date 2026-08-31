'use strict';

const state = {
  tariffs: [], tariffMap: new Map(), surveillanceMap: new Map(), dampingMap: new Map(),
  countries: {}, meta: {}, interestRates: [], lastResult: null, bulkRows: [], filteredTariffs: [],
  tariffPage: 1, tariffFilter: 'all', pageSize: 50, suggestionIndex: -1
};

const $ = (id) => document.getElementById(id);
const qsa = (s, root=document) => [...root.querySelectorAll(s)];
const toNumber = (v) => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  let s = String(v ?? '').trim().replace(/\s/g,'');
  if (s.includes(',') && s.includes('.')) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g,'').replace(',','.');
    else s = s.replace(/,/g,'');
  } else if (s.includes(',')) s = s.replace(',','.');
  const n = Number(s); return Number.isFinite(n) ? n : 0;
};
const trUpper = (v) => String(v ?? '').replace(/\u00a0/g,' ').trim().replace(/\s+/g,' ').toLocaleUpperCase('tr-TR');
const escapeHtml = (v) => String(v ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
const normalizeGtip = (value) => {
  const raw = String(value ?? '').trim();
  const digits = raw.replace(/\D/g,'');
  if (digits.length === 12) return `${digits.slice(0,4)}.${digits.slice(4,6)}.${digits.slice(6,8)}.${digits.slice(8,10)}.${digits.slice(10,12)}`;
  return raw;
};
const fmt = (n, digits=2) => new Intl.NumberFormat('tr-TR',{minimumFractionDigits:digits,maximumFractionDigits:digits}).format(toNumber(n));
const fmtRate = (n) => new Intl.NumberFormat('tr-TR',{maximumFractionDigits:4}).format(toNumber(n));
const fmtPercent = (rawRate) => `%${fmtRate(toNumber(rawRate))}`;
const fmtBandrol = (rawRate) => `%${fmtRate(toNumber(rawRate))}`;
const DATA_SCHEMA_VERSION = 7;
const formatMoney = (n, label='TL') => `${fmt(n)} ${label}`;
const unique = (arr) => [...new Set(arr.filter(Boolean))];

let xlsxLoaderPromise = null;
function ensureXlsx(){
  if(window.XLSX) return Promise.resolve(window.XLSX);
  if(xlsxLoaderPromise) return xlsxLoaderPromise;
  xlsxLoaderPromise = new Promise((resolve,reject)=>{
    const script=document.createElement('script');
    script.src='https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
    script.async=true;
    const timer=setTimeout(()=>reject(new Error('Excel okuma bileşeni zaman aşımına uğradı. Dosyayı CSV olarak kaydedip tekrar deneyin.')),12000);
    script.onload=()=>{clearTimeout(timer);window.XLSX?resolve(window.XLSX):reject(new Error('Excel okuma bileşeni başlatılamadı.'));};
    script.onerror=()=>{clearTimeout(timer);reject(new Error('Excel okuma bileşeni yüklenemedi. Dosyayı CSV olarak kaydedip tekrar deneyin.'));};
    document.head.appendChild(script);
  });
  return xlsxLoaderPromise;
}

function toast(message){
  const el=$('toast'); el.textContent=message; el.classList.add('show');
  clearTimeout(toast.timer); toast.timer=setTimeout(()=>el.classList.remove('show'),2600);
}

async function loadJson(path){
  const res = await fetch(path,{cache:'no-store'});
  if(!res.ok) throw new Error(`${path} yüklenemedi (${res.status})`);
  return res.json();
}

async function getDataPackage(){
  if(window.VERGI_DATA && typeof window.VERGI_DATA === 'object') return window.VERGI_DATA;
  if(window.VERGI_DATA_EMBEDDED && typeof window.VERGI_DATA_EMBEDDED === 'object') return window.VERGI_DATA_EMBEDDED;
  throw new Error('vergi-data.js yüklenemedi veya tarayıcı eski dosyayı önbellekten kullandı.');
}

async function init(){
  try{
    $('loadingText').textContent='Tarifeler ve hesaplama kuralları okunuyor…';
    const dataPackage = await getDataPackage();
    const {tariffs=[], surveillance=[], damping=[], countries={}, meta={}, interestRates=[]} = dataPackage;
    state.interestRates=interestRates;
    if(Number(meta.dataSchemaVersion)!==DATA_SCHEMA_VERSION || meta.rateStorage!=='excel_raw') throw new Error(`Hesaplama motoru ile veri dosyası uyumsuz (motor: ${DATA_SCHEMA_VERSION}, veri: ${meta.dataSchemaVersion ?? 'yok'}).`);
    state.tariffs=tariffs; state.countries=countries; state.meta=meta;
    tariffs.forEach(t=>state.tariffMap.set(t.gtip,t));
    surveillance.forEach(g=>state.surveillanceMap.set(g.gtip,g));
    damping.forEach(d=>{
      const key=`${d.gtip}|${trUpper(d.mense)}`;
      if(!state.dampingMap.has(key)) state.dampingMap.set(key,[]);
      state.dampingMap.get(key).push(d);
    });
    state.abSet=new Set((countries.ab||[]).map(trUpper));
    state.aliases=new Map(Object.entries(countries.aliases||{}).map(([a,b])=>[trUpper(a),trUpper(b)]));
    state.filteredTariffs=tariffs;
    populateLists(); bindEvents(); updateMeta(); renderTariffTable(); renderHistory(); renderProducts(); renderDashboard();
    $('loadingText').textContent='Hazır';
    setTimeout(()=>$('loadingScreen').classList.add('done'),250);
  } catch(err){
    console.error(err);
    $('loadingText').textContent=`Vergi verisi okunamadı: ${err?.message || 'Bilinmeyen yükleme hatası'}`;
    document.querySelector('.loader-mark').textContent='!';
  }
}

function populateLists(){
  const countries=unique([...(state.countries.all||[]),...Object.values(state.countries.aliases||{})]).sort((a,b)=>a.localeCompare(b,'tr'));
  $('countryList').innerHTML=countries.map(c=>`<option value="${escapeHtml(c)}"></option>`).join('');
  const payments=unique(['PEŞİN','MAL MUKABİLİ',...(state.countries.payments||[]).map(trUpper)]);
  const agreements=unique(['DI','AT',...(state.countries.agreements||[]).map(trUpper)]).filter(a=>trUpper(a)!=='A.TR');
  $('paymentMethod').innerHTML='<option value="">Ödeme şekli seçin</option>'+payments.map(p=>`<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join('');
  $('agreement').innerHTML='<option value="">Uluslararası antlaşma seçin</option>'+agreements.map(a=>`<option value="${escapeHtml(a)}">${escapeHtml(a)}</option>`).join('');
  $('paymentMethod').value=''; $('agreement').value='';
  for(const id of ['scenarioAPayment','scenarioBPayment']){const el=$(id);if(el)el.innerHTML=payments.map(p=>`<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join('');}
}


function updateMeta(){
  const m=state.meta;
  const val=(id,n)=>{const el=$(id);if(el)el.textContent=new Intl.NumberFormat('tr-TR').format(n||0);};
  val('heroTariffCount',m.tariffCount); val('heroDampingCount',m.dampingCount); val('heroAliasCount',m.aliasCount);
  val('metaTariffs',m.tariffCount); val('metaSurveillance',m.surveillanceCount); val('metaDamping',m.dampingCount); val('metaAliases',m.aliasCount);
  const sourceName=$('sourceName'), sourceVersion=$('sourceVersion');
  if(sourceName)sourceName.textContent=m.source||'—';
  if(sourceVersion)sourceVersion.textContent=`${m.version||''} · ${m.generated||''}`;
}

function bindEvents(){
  const on=(id,event,handler)=>{const el=$(id);if(el)el.addEventListener(event,handler);};
  qsa('.nav-link').forEach(btn=>btn.addEventListener('click',()=>switchView(btn.dataset.view)));
  on('menuButton','click',()=>{const nav=$('mainNav');nav.classList.toggle('open');$('menuButton').setAttribute('aria-expanded',nav.classList.contains('open'));});
  on('calcForm','submit',onCalculate);
  on('calcForm','reset',()=>setTimeout(resetResult,0));
  on('gtipInput','input',onGtipInput);
  on('gtipInput','keydown',onSuggestionKeydown);
  on('gtipInput','blur',()=>setTimeout(()=>hideSuggestions(),160));
  on('gtipSearchButton','click',()=>{onGtipInput();$('gtipInput').focus();});
  on('downloadResult','click',downloadSingleResult);
  on('printResult','click',()=>window.print());
  on('tariffSearch','input',()=>{state.tariffPage=1;applyTariffFilter();});
  qsa('.filter-chip').forEach(b=>b.addEventListener('click',()=>{qsa('.filter-chip').forEach(x=>x.classList.remove('active'));b.classList.add('active');state.tariffFilter=b.dataset.filter;state.tariffPage=1;applyTariffFilter();}));
  on('prevPage','click',()=>{state.tariffPage--;renderTariffTable();});
  on('nextPage','click',()=>{state.tariffPage++;renderTariffTable();});
  on('tariffTableBody','click',e=>{const b=e.target.closest('[data-gtip]');if(b)openTariffModal(b.dataset.gtip);});
  on('closeModal','click',()=>$('tariffModal').close());
  on('downloadTemplate','click',downloadTemplate);
  const file=$('bulkFile'), zone=$('dropZone');
  if(file)file.addEventListener('change',()=>file.files[0]&&processBulkFile(file.files[0]));
  if(zone){
    ['dragenter','dragover'].forEach(evt=>zone.addEventListener(evt,e=>{e.preventDefault();zone.classList.add('drag');}));
    ['dragleave','drop'].forEach(evt=>zone.addEventListener(evt,e=>{e.preventDefault();zone.classList.remove('drag');}));
    zone.addEventListener('drop',e=>{const f=e.dataTransfer.files[0];if(f)processBulkFile(f);});
  }
  on('exportBulk','click',exportBulk);
  on('currencyCode','change',()=>{if($('currencyCode').value==='TL')$('exchangeRate').value='1';});
  on('interestForm','submit',calculateInterestForm);
  document.addEventListener('click',e=>{if(!e.target.closest('.gtip-field'))hideSuggestions();});
}

function switchView(name){
  qsa('.view').forEach(v=>v.classList.toggle('active',v.id===`view-${name}`));
  qsa('.nav-link').forEach(b=>b.classList.toggle('active',b.dataset.view===name));
  $('mainNav').classList.remove('open'); if(name==='urunler')renderProducts(); if(name==='rapor')renderDashboard(); window.scrollTo({top:0,behavior:'smooth'});
}

function standardCountry(raw){
  const n=trUpper(raw); return state.aliases.get(n)||n;
}

function calculate(input){
  const gtip=normalizeGtip(input.gtip), tariff=state.tariffMap.get(gtip);
  if(!tariff) return {error:`${gtip||'Girilen tarife'} 2026 vergi tablosunda bulunamadı.`,gtip};
  const rawInvoice=String(input.invoiceValue ?? '').trim(), rawFreight=String(input.freightCost ?? '').trim(), rawInsurance=String(input.insuranceCost ?? '').trim(), rawExtra=String(input.extraCost ?? '').trim();
  const rawQuantity=String(input.quantity ?? '').trim(), rawGross=String(input.grossWeight ?? '').trim(), rawNet=String(input.netWeight ?? '').trim();
  const currencyCode=trUpper(input.currencyCode||'TL'), exchangeRate=currencyCode==='TL'?1:toNumber(input.exchangeRate), dollarRate=toNumber(input.dollarRate);
  const invoiceForeign=toNumber(rawInvoice), freightForeign=toNumber(rawFreight), insuranceForeign=toNumber(rawInsurance), otherExtraForeign=toNumber(rawExtra);
  // Döviz kuru yalnızca Mal Bedeline uygulanır. Yan masraflar kullanıcı tarafından TL girilir.
  const invoice=invoiceForeign*exchangeRate, freight=freightForeign, insurance=insuranceForeign, otherExtra=otherExtraForeign;
  const extra=freight+insurance+otherExtra, quantity=toNumber(rawQuantity), gross=toNumber(rawGross), net=toNumber(rawNet);
  const origin=standardCountry(input.originCountry), payment=trUpper(input.paymentMethod), agreement=trUpper(input.agreement), deliveryTerm=trUpper(input.deliveryTerm), quantityUnit=input.quantityUnit||'';
  if(!rawInvoice || invoiceForeign<=0) return {error:'Mal bedeli girilmeden hesaplama yapılamaz.',gtip,origin,agreement};
  if(currencyCode!=='TL' && exchangeRate<=0) return {error:'Dövizli işlemde geçerli bir kur girilmelidir.',gtip,origin,agreement};
  if(rawFreight===''||rawInsurance===''||rawExtra==='') return {error:'Navlun, sigorta ve diğer yan masraf alanlarına tutar veya 0 girilmelidir.',gtip,origin,agreement};
  if(!deliveryTerm) return {error:'Teslim şekli seçilmeden hesaplama yapılamaz.',gtip,origin,agreement};
  if(!origin) return {error:'Menşei girilmeden hesaplama yapılamaz.',gtip,origin,agreement};
  if(!rawQuantity || quantity<=0) return {error:'Miktar girilmeden hesaplama yapılamaz.',gtip,origin,agreement};
  if(!rawGross || gross<=0) return {error:'Brüt KG girilmeden hesaplama yapılamaz.',gtip,origin,agreement};
  if(!rawNet || net<=0) return {error:'Net KG girilmeden hesaplama yapılamaz.',gtip,origin,agreement};
  if(!payment) return {error:'Ödeme şekli seçilmeden hesaplama yapılamaz.',gtip,origin,agreement};
  if(!['DI','AT','A.TR'].includes(agreement)) return {error:'Geçerli uluslararası antlaşma bilgisi olmadan doğru vergi hesaplanamaz.',gtip,origin,agreement};
  const currency='TL', notices=[];
  // Gümrük kıymetinin başlangıç noktası: Mal Bedeli × kendi döviz kuru.
  // Navlun, sigorta ve diğer yan masraflar TL olarak ayrıca eklenir; döviz kuru bu kalemlere uygulanmaz.
  const baseCif=invoice+extra;
  const g=state.surveillanceMap.get(gtip)||{oran:0,birim:'',aciklama:''};
  let surveillance=0, surveillanceReason='GTİP için aktif gözetim kaydı bulunmadı.';
  let surveillanceTargetUSD=0, surveillanceTargetTL=0, priceCompletionUSD=0, surveillanceNewInvoiceTL=invoice;
  if(toNumber(g.oran)>0){
    if(dollarRate<=0) return {error:'Bu GTİP gözetim kapsamında. Gözetim fiyat tamamlama hesabı için geçerli Dolar Kuru girilmelidir.',gtip,origin,agreement};
    surveillanceTargetUSD=trUpper(g.birim)==='ADET' ? toNumber(g.oran)*quantity : gross*toNumber(g.oran);
    surveillanceTargetTL=surveillanceTargetUSD*dollarRate;
    // Fiyat tamamlama önce USD karşılığı üzerinden bulunur, ardından Dolar Kuru ile TL'ye çevrilir.
    const invoiceUsdEquivalent=invoice/dollarRate;
    priceCompletionUSD=Math.max(0,surveillanceTargetUSD-invoiceUsdEquivalent);
    surveillance=priceCompletionUSD*dollarRate;
    surveillanceNewInvoiceTL=invoice+surveillance;
    surveillanceReason=surveillance>0?'Gözetim sonrası yeni kıymet, Mal Bedelinin TL karşılığından yüksek olduğu için fiyat tamamlama oluştu.':'Mal Bedelinin TL karşılığı gözetim kıymetini karşıladığı için fiyat tamamlama oluşmadı.';
  }
  const cif=baseCif+surveillance;
  const gv=agreement==='DI' ? cif*toNumber(tariff.gv)/100 : 0;
  const gvReason=agreement==='DI'?(toNumber(tariff.gv)>0?'Dİ seçildi; GTİP satırındaki GV oranı uygulandı.':'Dİ seçildi ancak GTİP satırındaki GV oranı 0.'):`${agreement} seçildiği için Excel formülüne göre GV hesaplanmadı.`;
  const igvExempt=agreement==='AT'&&state.abSet.has(origin);
  const igv=igvExempt ? 0 : cif*toNumber(tariff.igv)/100;
  const igvReason=igvExempt?'AB menşeli ürün ve AT seçildiği için İGV hesaplanmadı.':(toNumber(tariff.igv)>0?'GTİP satırındaki İGV oranı uygulandı.':'GTİP satırındaki İGV oranı 0.');
  let gtsRate=0, emyReason='Belge, GV veya menşe koşulu sağlanmadığı için EMY-GTS hesaplanmadı.';
  if((agreement==='AT'||agreement==='A.TR')&&gv===0){
    const rates=tariff.gts||{};
    if(origin==='BİLİNMEYEN ÜLKE TOPRAKLARI'||origin==='BILINMEYEN ÜLKE TOPRAKLARI') gtsRate=Math.max(0,...Object.values(rates).map(toNumber));
    else gtsRate=toNumber(rates[origin]);
    emyReason=gtsRate>0?'Menşe ve GTİP için EMY-GTS oran satırı bulundu.':'Seçilen menşe için EMY-GTS oranı bulunmadı.';
  }
  const emy=cif*gtsRate/100;
  const bandrol=cif*toNumber(tariff.bandrol)/100;
  const bandrolReason=toNumber(tariff.bandrol)>0?'Excel V10 bandrol formülündeki oran CIF matrahına /100 uygulanarak hesaplandı.':'GTİP satırındaki bandrol oranı 0.';
  const otvBase=cif+gv+igv+emy+bandrol, otv=otvBase*toNumber(tariff.otv)/100;
  const otvReason=toNumber(tariff.otv)>0?'ÖTV matrahına CIF, GV, İGV, EMY-GTS ve bandrol eklendi.':'GTİP satırındaki ÖTV oranı 0.';
  const kkdf=payment==='MAL MUKABİLİ'?cif*toNumber(tariff.kkdf)/100:0;
  const kkdfReason=payment==='MAL MUKABİLİ'?(toNumber(tariff.kkdf)>0?'Mal mukabili ödeme seçildiği için KKDF uygulandı.':'Mal mukabili seçildi ancak KKDF oranı 0.'):'Ödeme şekli Mal Mukabili olmadığı için KKDF hesaplanmadı.';
  const dampingCandidates=state.dampingMap.get(`${gtip}|${origin}`)||[], dampingRec=dampingCandidates[0]||null;
  let dampingRate=0,damping=0,dampingFormula='Uygulanmadı',dampingBase=0,dampingReason='GTİP ve menşe için damping kaydı bulunmadı.';
  if(dampingRec){
    dampingRate=toNumber(dampingRec.ust); const unit=trUpper(dampingRec.birim);
    if(unit==='CIF'){dampingBase=cif;damping=cif*dampingRate/100;dampingFormula=`${fmt(cif)} × ${fmtRate(dampingRate)} / 100`;}
    else if(unit==='ADET'){dampingBase=quantity;damping=Math.max(0,dampingRate*quantity-cif);dampingFormula=`MAK(0; ${fmtRate(dampingRate)} × ${fmt(quantity,3)} − ${fmt(cif)})`;}
    else if(unit==='KG'){dampingBase=gross;damping=Math.max(0,dampingRate*gross-cif);dampingFormula=`MAK(0; ${fmtRate(dampingRate)} × ${fmt(gross,3)} − ${fmt(cif)})`;}
    dampingReason=`GTİP ve menşe için damping kaydı bulundu; Excel formülündeki üst değer kullanıldı (${dampingRec.metin||'kaynak kayıt'}).`;
  }
  const kdvBase=cif+gv+igv+emy+bandrol+otv+kkdf+damping, kdv=kdvBase*toNumber(tariff.kdv)/100;
  const kdvReason=toNumber(tariff.kdv)>0?'KDV matrahına CIF, GV, İGV, EMY-GTS, bandrol, ÖTV, KKDF ve damping eklendi.':'GTİP satırındaki KDV oranı 0.';
  const total=surveillance+gv+igv+emy+bandrol+otv+kkdf+damping+kdv, landedCost=baseCif+total;
  if(origin!==trUpper(input.originCountry)) notices.push({type:'info',text:`Menşe “${input.originCountry}” → “${origin}” olarak standartlaştırıldı.`});
  if(gross<net) notices.push({type:'warn',text:'Brüt KG, net KG değerinden küçük. Ağırlıkları kontrol edin.'});
  if(Object.keys(tariff.ek||{}).length) notices.push({type:'warn',text:'Bu GTİP için Excel tablosunda ek mevzuat/açıklama kayıtları var; GTİP detayını kontrol edin.'});
  if(tariff.kontrol?.gv||tariff.kontrol?.kdv||tariff.kontrol?.otv||tariff.kontrol?.bandrol) notices.push({type:'warn',text:'Kaynak satırda şartlı veya metin içeren vergi kaydı bulunuyor; uzman kontrolü önerilir.'});
  const src=state.meta.source||'2026 Vergi Tablosu';
  const lines=[
    {name:'Gözetim Fiyat Tamamlama Tutarı',rate:g.oran?`${fmtRate(g.oran)} USD / ${g.birim||'birim'}`:'—',base:invoice,formula:g.oran?`MAK(0; (${fmtRate(g.oran)} × ${trUpper(g.birim)==='ADET'?fmt(quantity,3):fmt(gross,3)}) − (${fmt(invoice)} / ${fmt(dollarRate,6)})) × ${fmt(dollarRate,6)}`:'Uygulanmadı',reason:surveillanceReason,source:'Gözetim sayfası',amount:surveillance},
    {name:'Gümrük Vergisi (GV)',rate:fmtPercent(tariff.gv),base:cif,formula:gv?`${fmt(cif)} × ${fmtRate(tariff.gv)} / 100`:'Uygulanmadı',reason:gvReason,source:src,amount:gv},
    {name:'İlave Gümrük Vergisi (İGV)',rate:fmtPercent(tariff.igv),base:cif,formula:igv?`${fmt(cif)} × ${fmtRate(tariff.igv)} / 100`:'Uygulanmadı',reason:igvReason,source:src,amount:igv},
    {name:'Ek Mali Yükümlülük (GTS)',rate:fmtPercent(gtsRate),base:cif,formula:emy?`${fmt(cif)} × ${fmtRate(gtsRate)} / 100`:'Uygulanmadı',reason:emyReason,source:'2026 Vergi Tablosu · Menşe sütunları',amount:emy},
    {name:'Bandrol',rate:tariff.bandrol?fmtBandrol(tariff.bandrol):'—',base:cif,formula:bandrol?`${fmt(cif)} × ${fmtRate(tariff.bandrol)} / 100`:'Uygulanmadı',reason:bandrolReason,source:src,amount:bandrol},
    {name:'Özel Tüketim Vergisi (ÖTV)',rate:fmtPercent(tariff.otv),base:otvBase,formula:otv?`${fmt(otvBase)} × ${fmtRate(tariff.otv)} / 100`:'Uygulanmadı',reason:otvReason,source:src,amount:otv},
    {name:'KKDF',rate:fmtPercent(tariff.kkdf),base:cif,formula:kkdf?`${fmt(cif)} × ${fmtRate(tariff.kkdf)} / 100`:'Uygulanmadı',reason:kkdfReason,source:src,amount:kkdf},
    {name:'Damping',rate:dampingRec?(dampingRec.para==='%'?fmtPercent(dampingRate):`${fmtRate(dampingRate)} ${dampingRec.para||''}`):'—',base:dampingBase,formula:dampingFormula,reason:dampingReason,source:'Damping sayfası',amount:damping},
    {name:'KDV',rate:fmtPercent(tariff.kdv),base:kdvBase,formula:kdv?`${fmt(kdvBase)} × ${fmtRate(tariff.kdv)} / 100`:'Uygulanmadı',reason:kdvReason,source:src,amount:kdv}
  ];
  return {gtip,tariff,origin,payment,agreement,deliveryTerm,quantityUnit,currency,currencyCode,exchangeRate,dollarRate,invoiceForeign,freightForeign,insuranceForeign,otherExtraForeign,invoice,freight,insurance,otherExtra,extra,quantity,gross,net,surveillance,surveillanceTargetUSD,surveillanceTargetTL,priceCompletionUSD,surveillanceNewInvoiceTL,surveillanceRecord:g,dampingRec,dampingRate,cif,gv,igv,gtsRate,emy,bandrol,otv,kkdf,damping,kdv,total,landedCost,notices,lines,materialCode:input.materialCode||'',commercialDescription:input.commercialDescription||'',timestamp:new Date().toISOString()};
}
function collectCurrentInput(){return {
  gtip:$('gtipInput').value,currencyCode:$('currencyCode').value,exchangeRate:$('exchangeRate').value,dollarRate:$('dollarRate').value,invoiceValue:$('invoiceValue').value,deliveryTerm:$('deliveryTerm').value,
  freightCost:$('freightCost').value,insuranceCost:$('insuranceCost').value,extraCost:$('extraCost').value,quantity:$('quantity').value,quantityUnit:'',
  grossWeight:$('grossWeight').value,netWeight:$('netWeight').value,originCountry:$('originCountry').value,paymentMethod:$('paymentMethod').value,agreement:$('agreement').value,
  dampingChoice:$('dampingChoice').value,materialCode:$('materialCode').value,commercialDescription:$('commercialDescription').value};}
function onCalculate(e){
  e.preventDefault(); const result=calculate(collectCurrentInput());
  if(result.error){$('gtipStatus').textContent=result.error;$('gtipStatus').className='field-help error';toast(result.error);return;}
  state.lastResult=result; renderResult(result); saveHistory(result);
}
function resultZeroNote(line){
  if(Math.abs(toNumber(line.amount))>0.000001) return '—';
  const name=trUpper(line.name||''), reason=trUpper(line.reason||'');
  if(name.includes('GÖZETİM')) return reason.includes('AKTİF GÖZETİM KAYDI BULUNMADI')?'Gözetim kaydı bulunmuyor.':'Fiyat tamamlama oluşmadı.';
  if(name.includes('GÜMRÜK VERGİSİ')) return reason.includes('SEÇİLDİĞİ İÇİN')?'Belge seçimi nedeniyle GV uygulanmadı.':'GV oranı %0.';
  if(name.includes('İLAVE GÜMRÜK')) return reason.includes('AB MENŞELİ')?'AB menşe + AT nedeniyle İGV uygulanmadı.':'İGV oranı %0.';
  if(name.includes('EK MALİ')) return reason.includes('ORANI BULUNMADI')?'Menşe için EMY-GTS oranı yok.':'EMY-GTS koşulları sağlanmadı.';
  if(name==='BANDROL') return 'Bandrol oranı yok / %0.';
  if(name.includes('ÖZEL TÜKETİM')) return 'ÖTV oranı %0.';
  if(name==='KKDF') return reason.includes('ÖDEME ŞEKLİ')?'Ödeme şekli nedeniyle KKDF uygulanmadı.':'KKDF oranı %0.';
  if(name==='DAMPİNG'||name==='DAMPING') return line.base>0?'Damping eşiği aşılmadı.':'GTİP/menşe için damping kaydı yok.';
  if(name==='KDV') return 'KDV oranı %0.';
  return line.reason?String(line.reason).replace(/\s+/g,' ').slice(0,90):'Uygulanmadı.';
}
function renderResult(r){
  $('emptyResult').hidden=true; $('resultContent').hidden=false;
  $('landedCost').textContent=formatMoney(r.landedCost,r.currency); $('totalTax').textContent=formatMoney(r.total,r.currency); $('cifValue').textContent=formatMoney(r.cif,r.currency); $('resultGtip').textContent=r.gtip;
  $('resultNotices').innerHTML=r.notices.map(n=>`<div class="notice ${n.type==='info'?'':n.type}">${escapeHtml(n.text)}</div>`).join('');
  const pills=[`GV ${fmtPercent(r.tariff.gv)}`,`İGV ${fmtPercent(r.tariff.igv)}`,`KDV ${fmtPercent(r.tariff.kdv)}`,`ÖTV ${fmtPercent(r.tariff.otv)}`,`KKDF ${fmtPercent(r.tariff.kkdf)}`];
  if(r.gtsRate)pills.push(`EMY-GTS ${fmtPercent(r.gtsRate)}`); if(r.tariff.bandrol)pills.push(`Bandrol ${fmtPercent(r.tariff.bandrol)}`);
  $('rateStrip').innerHTML=pills.map(x=>`<span class="rate-pill">${escapeHtml(x)}</span>`).join('');
  $('resultRows').innerHTML=r.lines.map(x=>`<tr><td><strong>${escapeHtml(x.name)}</strong></td><td>${escapeHtml(x.rate)}</td><td>${formatMoney(x.base||0,r.currency)}</td><td><strong>${formatMoney(x.amount,r.currency)}</strong></td><td class="result-note">${escapeHtml(resultZeroNote(x))}</td></tr>`).join('');
  const details={'Eşya tanımı':r.tariff.tanim||'—','Ticari tanım':r.commercialDescription||'—','Malzeme kodu':r.materialCode||'—','Menşe':r.origin,'Belge / anlaşma':r.agreement,'Ödeme şekli':r.payment,'Teslim şekli':r.deliveryTerm,'Mal Bedeli döviz / kur':`${r.currencyCode} · ${fmt(r.exchangeRate,6)} (yalnızca mal bedeline uygulanır)`,'Dolar Kuru':r.dollarRate>0?`${fmt(r.dollarRate,6)} TL/USD`:'—','Mal bedeli (TL)':formatMoney(r.invoice,r.currency),'Gözetim sonrası mal bedeli (TL)':formatMoney(r.surveillanceNewInvoiceTL||r.invoice,r.currency),'Gözetim fiyat tamamlama (TL)':formatMoney(r.surveillance,r.currency),'Navlun + sigorta + diğer (TL)':formatMoney(r.extra,r.currency),'Brüt / net':`${fmt(r.gross,3)} / ${fmt(r.net,3)} kg`,'Miktar':fmt(r.quantity,3),'Gözetim açıklaması':r.surveillanceRecord.aciklama||'Yok','Damping kaydı':r.dampingRec?.metin||'Yok','KDV matrahı':formatMoney(r.lines.find(x=>x.name==='KDV')?.base||0,r.currency)};
  $('tariffDetail').innerHTML=Object.entries(details).map(([k,v])=>`<div class="detail-item"><span>${escapeHtml(k)}</span><strong>${escapeHtml(v)}</strong></div>`).join('');
  $('gtipStatus').textContent=`${r.gtip} bulundu ve hesaplandı.`;$('gtipStatus').className='field-help ok';
}
function resetResult(){state.lastResult=null;$('emptyResult').hidden=false;$('resultContent').hidden=true;$('gtipStatus').textContent='12 haneli GTİP girin.';$('gtipStatus').className='field-help';}

function onGtipInput(){
  const el=$('gtipInput'), value=el.value; const normalized=normalizeGtip(value);
  if(value.replace(/\D/g,'').length===12 && value!==normalized) el.value=normalized;
  const query=el.value.replace(/\D/g,'');
  if(query.length<2){hideSuggestions();return;}
  const matches=state.tariffs.filter(t=>t.gtip.replace(/\D/g,'').includes(query)).slice(0,9);
  if(!matches.length){hideSuggestions();return;}
  state.suggestionIndex=-1;
  $('gtipSuggestions').innerHTML=matches.map((t,i)=>`<button type="button" class="suggestion-item" data-index="${i}" data-value="${t.gtip}" role="option"><strong>${t.gtip}</strong><small>GV ${fmtPercent(t.gv)} · İGV ${fmtPercent(t.igv)} · KDV ${fmtPercent(t.kdv)}</small></button>`).join('');
  $('gtipSuggestions').hidden=false;
  qsa('.suggestion-item',$('gtipSuggestions')).forEach(b=>b.addEventListener('mousedown',e=>{e.preventDefault();selectSuggestion(b.dataset.value);}));
  const exact=state.tariffMap.get(normalized); $('gtipStatus').textContent=exact?'Tarife bulundu.':'Eşleşen tarifeler gösteriliyor.';$('gtipStatus').className=`field-help ${exact?'ok':''}`;
}
function selectSuggestion(value){$('gtipInput').value=value;hideSuggestions();$('gtipStatus').textContent='Tarife bulundu.';$('gtipStatus').className='field-help ok';}
function hideSuggestions(){$('gtipSuggestions').hidden=true;state.suggestionIndex=-1;}
function onSuggestionKeydown(e){const items=qsa('.suggestion-item',$('gtipSuggestions'));if($('gtipSuggestions').hidden||!items.length)return;if(e.key==='ArrowDown'){e.preventDefault();state.suggestionIndex=Math.min(items.length-1,state.suggestionIndex+1);}else if(e.key==='ArrowUp'){e.preventDefault();state.suggestionIndex=Math.max(0,state.suggestionIndex-1);}else if(e.key==='Enter'&&state.suggestionIndex>=0){e.preventDefault();selectSuggestion(items[state.suggestionIndex].dataset.value);return;}else if(e.key==='Escape'){hideSuggestions();return;}items.forEach((x,i)=>x.classList.toggle('active',i===state.suggestionIndex));items[state.suggestionIndex]?.scrollIntoView({block:'nearest'});}

function csvEscape(v){const s=String(v??'');return /[;"\n]/.test(s)?`"${s.replace(/"/g,'""')}"`:s;}
function downloadBlob(content,name,type='text/csv;charset=utf-8'){const blob=new Blob([content],{type});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);}
function downloadSingleResult(){if(!state.lastResult){toast('Önce hesaplama yapın.');return;}const r=state.lastResult;const rows=[['GTİP',r.gtip],['Menşe',r.origin],['Belge',r.agreement],['Ödeme Şekli',r.payment],['Teslim Şekli',r.deliveryTerm],['CIF Kıymeti',r.cif],...r.lines.map(x=>[x.name,x.rate,x.base,x.amount,resultZeroNote(x)]),['Vergi + Tamamlama Toplamı',r.total],['Toplam İthalat Maliyeti',r.landedCost]];downloadBlob('\ufeff'+rows.map(x=>x.map(csvEscape).join(';')).join('\n'),`vergi_hesap_${r.gtip.replace(/\./g,'_')}.csv`);}
function applyTariffFilter(){
  const q=$('tariffSearch').value.replace(/\D/g,'');
  state.filteredTariffs=state.tariffs.filter(t=>{
    const match=!q||t.gtip.replace(/\D/g,'').includes(q);
    if(!match)return false;
    if(state.tariffFilter==='igv')return toNumber(t.igv)>0;
    if(state.tariffFilter==='gts')return Object.keys(t.gts||{}).length>0;
    if(state.tariffFilter==='gozetim')return toNumber(state.surveillanceMap.get(t.gtip)?.oran)>0;
    if(state.tariffFilter==='bandrol')return toNumber(t.bandrol)>0;
    if(state.tariffFilter==='otv')return toNumber(t.otv)>0;
    return true;
  });renderTariffTable();
}
function renderTariffTable(){
  const total=Math.max(1,Math.ceil(state.filteredTariffs.length/state.pageSize));state.tariffPage=Math.min(Math.max(1,state.tariffPage),total);
  const rows=state.filteredTariffs.slice((state.tariffPage-1)*state.pageSize,state.tariffPage*state.pageSize);
  $('tariffTableBody').innerHTML=rows.map(t=>{const g=state.surveillanceMap.get(t.gtip);const gozetim=g&&toNumber(g.oran)>0?`${fmtRate(g.oran)} USD / ${escapeHtml(g.birim||'birim')}`:'—';return `<tr><td><strong>${t.gtip}</strong></td><td>${fmtPercent(t.gv)}</td><td>${fmtPercent(t.igv)}</td><td>${fmtPercent(t.kdv)}</td><td>${fmtPercent(t.otv)}</td><td>${t.bandrol?fmtBandrol(t.bandrol):'—'}</td><td>${fmtPercent(t.kkdf)}</td><td>${Object.keys(t.gts||{}).length?Object.keys(t.gts).length+' ülke':'—'}</td><td>${gozetim}</td><td><button class="link-button" data-gtip="${t.gtip}">Detay</button></td></tr>`;}).join('')||'<tr><td colspan="10">Eşleşen tarife bulunamadı.</td></tr>';
  $('pageInfo').textContent=`${state.tariffPage} / ${total} · ${new Intl.NumberFormat('tr-TR').format(state.filteredTariffs.length)} kayıt`;
  $('prevPage').disabled=state.tariffPage<=1;$('nextPage').disabled=state.tariffPage>=total;
}
function openTariffModal(gtip){
  const t=state.tariffMap.get(gtip),g=state.surveillanceMap.get(gtip);if(!t)return;
  $('modalTitle').textContent=gtip;
  const rates={'GV':fmtPercent(t.gv),'İGV':fmtPercent(t.igv),'KDV':fmtPercent(t.kdv),'ÖTV':fmtPercent(t.otv),'Bandrol':t.bandrol?fmtBandrol(t.bandrol):'—','KKDF':fmtPercent(t.kkdf)};
  let html=`<section class="modal-section"><h3>${escapeHtml(t.tanim||'Tarife')}</h3><div class="modal-list">${Object.entries(rates).map(([k,v])=>`<div><span>${k}</span><strong>${v}</strong></div>`).join('')}</div></section>`;
  const rawRates={'GV kaynak':t.gvMetin,'KDV kaynak':t.kdvMetin,'ÖTV kaynak':t.otvMetin,'İGV kaynak':t.igvMetin,'Bandrol kaynak':t.bandrolMetin,'KKDF kaynak':t.kkdfMetin};
  html+=`<section class="modal-section"><h3>Kaynak tarife metinleri</h3><div class="extra-list">${Object.entries(rawRates).filter(([,v])=>v).map(([k,v])=>`<div><span>${escapeHtml(k)}</span><strong>${escapeHtml(v)}</strong></div>`).join('')||'<div><span>Bilgi</span><strong>Ek metin yok</strong></div>'}</div></section>`;
  if(g&&(g.oran||g.aciklama))html+=`<section class="modal-section"><h3>Gözetim</h3><div class="extra-list"><div><span>Açıklama</span><strong>${escapeHtml(g.aciklama||'—')}</strong></div><div><span>Oran / Birim</span><strong>${fmtRate(g.oran)} ${escapeHtml(g.birim||'')}</strong></div></div></section>`;
  if(Object.keys(t.gts||{}).length)html+=`<section class="modal-section"><h3>EMY-GTS ülke oranları</h3><div class="extra-list">${Object.entries(t.gts).map(([k,v])=>`<div><span>${escapeHtml(k)}</span><strong>${fmtPercent(v)}</strong></div>`).join('')}</div></section>`;
  if(Object.keys(t.ek||{}).length)html+=`<section class="modal-section"><h3>Excel tablosundaki ek bilgiler</h3><div class="extra-list">${Object.entries(t.ek).map(([k,v])=>`<div><span>${escapeHtml(k)}</span><strong>${escapeHtml(v)}</strong></div>`).join('')}</div></section>`;
  const dampingCount=[...state.dampingMap.entries()].filter(([k])=>k.startsWith(gtip+'|')).reduce((n,[,v])=>n+v.length,0);
  html+=`<section class="modal-section"><h3>Kapsam özeti</h3><div class="modal-list"><div><span>Damping kaydı</span><strong>${dampingCount}</strong></div><div><span>Gözetim</span><strong>${g?.oran?'Var':'Yok'}</strong></div><div><span>Ek kayıt</span><strong>${Object.keys(t.ek||{}).length}</strong></div></div></section>`;
  $('modalBody').innerHTML=html;$('tariffModal').showModal();
}

function getHistory(){try{return JSON.parse(localStorage.getItem('vergiHesapHistory')||'[]');}catch{return[];}}
function saveHistory(r){const h=getHistory();h.unshift({gtip:r.gtip,origin:r.origin,agreement:r.agreement,payment:r.payment,deliveryTerm:r.deliveryTerm,currencyCode:r.currencyCode,exchangeRate:r.exchangeRate,dollarRate:r.dollarRate,invoiceForeign:r.invoiceForeign,freightForeign:r.freightForeign,insuranceForeign:r.insuranceForeign,otherExtraForeign:r.otherExtraForeign,invoice:r.invoice,extra:r.extra,quantity:r.quantity,quantityUnit:r.quantityUnit,gross:r.gross,net:r.net,total:r.total,landedCost:r.landedCost,currency:r.currency,lines:r.lines.map(x=>({name:x.name,amount:x.amount})),materialCode:r.materialCode,commercialDescription:r.commercialDescription,timestamp:r.timestamp});localStorage.setItem('vergiHesapHistory',JSON.stringify(h.slice(0,200)));renderHistory();renderDashboard();}
function renderHistory(){const el=$('historyGrid');if(!el)return;const h=getHistory();el.innerHTML=h.length?h.map((x,i)=>`<article class="history-card"><h3>${escapeHtml(x.materialCode||x.gtip)}</h3><p>${escapeHtml(x.gtip)} · ${escapeHtml(x.origin)} · ${escapeHtml(x.agreement)} · ${new Date(x.timestamp).toLocaleString('tr-TR')}</p><span class="history-total">${formatMoney(x.landedCost||x.total,x.currency)}</span><div class="history-actions"><button data-action="load" data-index="${i}">Yükle</button><button data-action="delete" data-index="${i}">Sil</button></div></article>`).join(''):'<div class="empty-history">Henüz kaydedilmiş hesaplama yok.</div>';}
function onHistoryClick(e){const b=e.target.closest('button[data-action]');if(!b)return;const h=getHistory(),i=Number(b.dataset.index);if(b.dataset.action==='delete'){h.splice(i,1);localStorage.setItem('vergiHesapHistory',JSON.stringify(h));renderHistory();renderDashboard();return;}const x=h[i];$('gtipInput').value=x.gtip;$('originCountry').value=x.origin;$('agreement').value=x.agreement;$('paymentMethod').value=x.payment;$('deliveryTerm').value=x.deliveryTerm||'CIF';$('currencyCode').value=x.currencyCode||'TL';$('exchangeRate').value=x.exchangeRate||1;if($('dollarRate'))$('dollarRate').value=x.dollarRate||'';$('invoiceValue').value=x.invoiceForeign??x.invoice;$('freightCost').value=x.freightForeign??0;$('insuranceCost').value=x.insuranceForeign??0;$('extraCost').value=x.otherExtraForeign??x.extra??0;$('quantity').value=x.quantity;if($('quantityUnit'))$('quantityUnit').value=x.quantityUnit||'Adet';$('grossWeight').value=x.gross;$('netWeight').value=x.net||x.gross;$('materialCode').value=x.materialCode||'';$('commercialDescription').value=x.commercialDescription||'';switchView('hesaplama');toast('Geçmiş hesaplama forma yüklendi.');}
function downloadTemplate(){const headers=['Tarife','Mal Bedeli','Döviz','Kur','Dolar Kuru','Teslim Şekli','Navlun','Sigorta','Diğer Yan Masraf','Menşei','Miktar','Birim','Brüt KG','Net KG','Ödeme Şekli','ULUSLARARASI ANT','Malzeme Kodu','Ticari Tanım'];downloadBlob('\ufeff'+headers.map(csvEscape).join(';')+'\n','toplu_vergi_hesaplama_sablonu.csv');}

async function processBulkFile(file){
  $('bulkStatus').textContent=`${file.name} okunuyor…`;$('bulkResults').hidden=true;
  try{
    const ext=file.name.split('.').pop().toLowerCase();let rows;
    if(ext==='csv')rows=parseCsv(await file.text());
    else{
      await ensureXlsx();
      const data=await file.arrayBuffer();const wb=XLSX.read(data,{type:'array',cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];rows=XLSX.utils.sheet_to_json(ws,{defval:''});
    }
    if(!rows.length)throw new Error('Dosyada hesaplanacak satır bulunamadı.');
    const dampingChoice=$('bulkDamping').value;
    state.bulkRows=rows.map((row,index)=>bulkCalculateRow(row,index,dampingChoice));
    renderBulk();
  }catch(err){console.error(err);$('bulkStatus').textContent=err.message;$('bulkStatus').style.color='var(--danger)';}
}
function parseCsv(text){
  const clean=text.replace(/^\ufeff/,'');const delimiter=(clean.split('\n')[0].match(/;/g)||[]).length>=(clean.split('\n')[0].match(/,/g)||[]).length?';':',';
  const lines=[];let row=[],cell='',quoted=false;
  for(let i=0;i<clean.length;i++){const ch=clean[i],next=clean[i+1];if(ch==='"'&&quoted&&next==='"'){cell+='"';i++;}else if(ch==='"'){quoted=!quoted;}else if(ch===delimiter&&!quoted){row.push(cell);cell='';}else if((ch==='\n'||ch==='\r')&&!quoted){if(ch==='\r'&&next==='\n')i++;row.push(cell);cell='';if(row.some(x=>x!==''))lines.push(row);row=[];}else cell+=ch;}row.push(cell);if(row.some(x=>x!==''))lines.push(row);const headers=lines.shift().map(x=>x.trim());return lines.map(r=>Object.fromEntries(headers.map((h,i)=>[h,r[i]??''])));}
function normHeader(h){return trUpper(h).replace(/[İI]/g,'I').replace(/[^A-Z0-9ÇĞÖŞÜ]/g,'');}
function findValue(row,aliases){const entries=Object.entries(row);for(const alias of aliases){const target=normHeader(alias);const found=entries.find(([k])=>normHeader(k)===target);if(found)return found[1];}return '';}
function bulkCalculateRow(row,index,dampingChoice){
  const rowPayment=findValue(row,['Ödeme Şekli','Odeme Sekli']), rowAgreement=findValue(row,['ULUSLARARASI ANT','Uluslararası Ant','Uluslararası Anlaşma','AT/DI','AT DI','AT-DI','Anlaşma','Anlasma']);
  const input={gtip:findValue(row,['Tarife','GTİP','GTIP']),invoiceValue:findValue(row,['Mal Bedeli','Malbedeli','Tutar','Fatura Tutarı']),currencyCode:findValue(row,['Döviz','Döviz Cinsi','Para Birimi'])||'TL',exchangeRate:findValue(row,['Kur','Döviz Kuru'])||1,dollarRate:findValue(row,['Dolar Kuru','USD Kuru','USD/TL'])||0,deliveryTerm:findValue(row,['Teslim Şekli','Incoterm'])||'CIF',freightCost:findValue(row,['Navlun','Freight'])||0,insuranceCost:findValue(row,['Sigorta','Insurance'])||0,extraCost:findValue(row,['Diğer Yan Masraf','Yan Masraf','Tedarik Dışı Yan Masraf'])||0,grossWeight:findValue(row,['Brüt KG','Brüt Ağırlık','GW KG']),netWeight:findValue(row,['Net KG','Net Ağırlık','NW KG']),quantity:findValue(row,['Miktar','Adet']),quantityUnit:findValue(row,['Birim','Miktar Birimi'])||'Adet',originCountry:findValue(row,['Menşei','Menşe Ülkesi','Menşe','Mense']),paymentMethod:rowPayment,agreement:rowAgreement,dampingChoice,materialCode:findValue(row,['Malzeme Kodu','Material']),commercialDescription:findValue(row,['Ticari Tanım','Tanım'])};
  const result=calculate(input);const base={...row,'Satır':index+2,'Hesap GTİP':normalizeGtip(input.gtip),'Standart Menşe':result.origin||standardCountry(input.originCountry),'Hesap AT/DI':trUpper(input.agreement)};
  if(result.error)return {...base,'Hata':result.error};
  return {...base,'CIF Kıymeti':result.cif,'Gözetim Fiyat Tamamlama Tutarı':result.surveillance,'GV':result.gv,'İGV':result.igv,'EMY-GTS':result.emy,'Bandrol':result.bandrol,'ÖTV':result.otv,'KKDF':result.kkdf,'Damping':result.damping,'KDV':result.kdv,'Vergi + Tamamlama Toplamı':result.total,'Toplam İthalat Maliyeti':result.landedCost,'Hesaplama Nedeni':result.lines.map(x=>`${x.name}: ${x.reason}`).join(' | '),'Hata':result.notices.filter(n=>n.type==='warn').map(n=>n.text).join(' | ')};
}
function renderBulk(){
  const rows=state.bulkRows,errors=rows.filter(r=>r.Hata).length;$('bulkResults').hidden=false;$('bulkStatus').textContent='Dosya tarayıcı içinde hesaplandı.';$('bulkStatus').style.color='var(--brand)';
  $('bulkSummary').textContent=`${new Intl.NumberFormat('tr-TR').format(rows.length)} satır hesaplandı`;$('bulkErrorSummary').textContent=errors?`${errors} satırda uyarı veya hata var.`:'Hata bulunmadı.';
  const keys=unique(rows.flatMap(Object.keys));const preferred=['Satır','Tarife','Mal Bedeli','Döviz','Kur','Dolar Kuru','Teslim Şekli','Navlun','Sigorta','Diğer Yan Masraf','Menşei','Miktar','Birim','Brüt KG','Net KG','Ödeme Şekli','ULUSLARARASI ANT','Hesap GTİP','Standart Menşe','Hesap AT/DI','CIF Kıymeti','GV','İGV','EMY-GTS','Bandrol','ÖTV','KKDF','Damping','KDV','Vergi + Tamamlama Toplamı','Toplam İthalat Maliyeti','Hesaplama Nedeni','Hata'];const cols=[...preferred.filter(k=>keys.includes(k)),...keys.filter(k=>!preferred.includes(k))].slice(0,24);
  state.bulkColumns=cols;$('bulkHead').innerHTML='<tr>'+cols.map(c=>`<th>${escapeHtml(c)}</th>`).join('')+'</tr>';
  $('bulkBody').innerHTML=rows.slice(0,300).map(r=>'<tr>'+cols.map(c=>`<td>${typeof r[c]==='number'?fmt(r[c]):escapeHtml(r[c]??'')}</td>`).join('')+'</tr>').join('');
  if(rows.length>300)$('bulkStatus').textContent+=' Önizlemede ilk 300 satır gösteriliyor; indirilen dosyada tüm satırlar yer alır.';
}
async function exportBulk(){if(!state.bulkRows.length){toast('Önce dosya yükleyin.');return;}try{await ensureXlsx();const ws=XLSX.utils.json_to_sheet(state.bulkRows);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Hesap Sonuçları');XLSX.writeFile(wb,'toplu_vergi_hesaplama_sonuclari.xlsx');}catch(e){const keys=unique(state.bulkRows.flatMap(Object.keys));const csv='\ufeff'+[keys,...state.bulkRows.map(r=>keys.map(k=>r[k]??''))].map(r=>r.map(csvEscape).join(';')).join('\n');downloadBlob(csv,'toplu_vergi_hesaplama_sonuclari.csv');toast('Excel bileşeni yüklenemedi; sonuçlar CSV olarak indirildi.');}}


function getProducts(){try{return JSON.parse(localStorage.getItem('vergiProductCards')||'[]');}catch{return[];}}
function saveCurrentProduct(){if(!state.lastResult){toast('Önce bir hesaplama yapın.');return;}const r=state.lastResult,p=getProducts(),key=r.materialCode||r.gtip;const item={key,materialCode:r.materialCode,description:r.commercialDescription||r.tariff.tanim||'',gtip:r.gtip,origin:r.origin,unit:r.quantityUnit,lastCalculated:r.timestamp,agreement:r.agreement,payment:r.payment};const i=p.findIndex(x=>x.key===key);if(i>=0)p[i]=item;else p.unshift(item);localStorage.setItem('vergiProductCards',JSON.stringify(p.slice(0,500)));renderProducts();toast('Ürün kartı kaydedildi.');}
function renderProducts(){const el=$('productGrid');if(!el)return;const p=getProducts();el.innerHTML=p.length?p.map((x,i)=>`<article class="product-card"><span class="product-code">${escapeHtml(x.materialCode||'GTİP')}</span><h3>${escapeHtml(x.description||x.gtip)}</h3><p>${escapeHtml(x.gtip)} · ${escapeHtml(x.origin)} · ${escapeHtml(x.unit)}</p><small>Son hesaplama: ${x.lastCalculated?new Date(x.lastCalculated).toLocaleString('tr-TR'):'—'}</small><div class="history-actions"><button data-product-action="load" data-index="${i}">Forma yükle</button><button data-product-action="delete" data-index="${i}">Sil</button></div></article>`).join(''):'<div class="empty-history">Henüz ürün kartı kaydedilmedi. Tekli hesaplama sonucundaki yıldız düğmesini kullanın.</div>';}
function onProductClick(e){const b=e.target.closest('button[data-product-action]');if(!b)return;const p=getProducts(),i=Number(b.dataset.index);if(b.dataset.productAction==='delete'){p.splice(i,1);localStorage.setItem('vergiProductCards',JSON.stringify(p));renderProducts();return;}const x=p[i];$('gtipInput').value=x.gtip;$('originCountry').value=x.origin;if($('quantityUnit'))$('quantityUnit').value=x.unit||'Adet';$('agreement').value=x.agreement||'';$('paymentMethod').value=x.payment||'';$('materialCode').value=x.materialCode||'';$('commercialDescription').value=x.description||'';switchView('hesaplama');toast('Ürün kartı forma yüklendi.');}
function prepareScenarioDefaults(){const base=collectCurrentInput();$('scenarioAOrigin').value=base.originCountry||'ÇİN HALK CUMHUR.';$('scenarioBOrigin').value=state.abSet.has(standardCountry(base.originCountry))?base.originCountry:'ALMANYA';$('scenarioAAgreement').value=base.agreement||'DI';$('scenarioBAgreement').value='A.TR';$('scenarioAPayment').value=base.paymentMethod||'PEŞİN';$('scenarioBPayment').value=base.paymentMethod||'PEŞİN';toast('Mevcut form değerleri senaryolara aktarıldı.');}
function compareScenarios(){const base=collectCurrentInput(),a=calculate({...base,originCountry:$('scenarioAOrigin').value,agreement:$('scenarioAAgreement').value,paymentMethod:$('scenarioAPayment').value}),b=calculate({...base,originCountry:$('scenarioBOrigin').value,agreement:$('scenarioBAgreement').value,paymentMethod:$('scenarioBPayment').value});const el=$('scenarioResult');el.hidden=false;if(a.error||b.error){el.innerHTML=`<div class="notice warn">${escapeHtml(a.error||b.error)}</div>`;return;}const diff=a.landedCost-b.landedCost;el.innerHTML=`<div class="scenario-summary"><div><span>Senaryo A</span><strong>${formatMoney(a.landedCost,'TL')}</strong><small>${escapeHtml(a.origin)} · ${escapeHtml(a.agreement)} · ${escapeHtml(a.payment)}</small></div><div><span>Senaryo B</span><strong>${formatMoney(b.landedCost,'TL')}</strong><small>${escapeHtml(b.origin)} · ${escapeHtml(b.agreement)} · ${escapeHtml(b.payment)}</small></div><div class="accent"><span>Tahmini fark (A − B)</span><strong>${formatMoney(diff,'TL')}</strong><small>${diff>0?'Senaryo B daha düşük maliyetli':diff<0?'Senaryo A daha düşük maliyetli':'Maliyetler eşit'}</small></div></div><div class="table-wrap"><table class="result-table"><thead><tr><th>Kalem</th><th>Senaryo A</th><th>Senaryo B</th><th>Fark</th></tr></thead><tbody>${a.lines.map((x,i)=>`<tr><td>${escapeHtml(x.name)}</td><td>${formatMoney(x.amount,'TL')}</td><td>${formatMoney(b.lines[i].amount,'TL')}</td><td>${formatMoney(x.amount-b.lines[i].amount,'TL')}</td></tr>`).join('')}</tbody></table></div>`;}
function calculateInterestForm(e){e.preventDefault();const principal=toNumber($('interestPrincipal').value),start=new Date($('interestStart').value+'T00:00:00'),end=new Date($('interestEnd').value+'T00:00:00');if(!principal||isNaN(start)||isNaN(end)||end<=start){toast('Ana para ve tarih aralığını kontrol edin.');return;}const ms=86400000,details=[];let total=0,covered=0;for(const p of state.interestRates){const ps=new Date(p.start+'T00:00:00'),pe=new Date(p.end+'T00:00:00');const days=Math.max(0,(Math.min(end,pe)-Math.max(start,ps))/ms);if(days>0){const amount=principal*(days/30)*(toNumber(p.monthlyRate)/100);total+=amount;covered+=days;details.push({...p,days,amount});}}const totalDays=(end-start)/ms,uncovered=Math.max(0,totalDays-covered);$('interestEmpty').hidden=true;$('interestResult').hidden=false;$('interestResult').innerHTML=`<div class="result-summary"><div class="summary-card accent"><span>Faiz tutarı</span><strong>${formatMoney(total,'TL')}</strong></div><div class="summary-card"><span>Ana para + faiz</span><strong>${formatMoney(principal+total,'TL')}</strong></div><div class="summary-card"><span>Toplam gün</span><strong>${fmt(totalDays,0)}</strong></div><div class="summary-card"><span>Efektif oran</span><strong>%${fmt(principal?total/principal*100:0,4)}</strong></div></div>${uncovered?`<div class="notice warn">${fmt(uncovered,0)} gün faiz tablosunun kapsadığı dönem dışında kaldı ve hesaplanmadı.</div>`:''}<div class="table-wrap"><table class="result-table"><thead><tr><th>Dönem</th><th>Aylık oran</th><th>Gün</th><th>Formül</th><th>Faiz</th></tr></thead><tbody>${details.map(x=>`<tr><td>${x.start} – ${x.end}</td><td>%${fmtRate(x.monthlyRate)}</td><td>${fmt(x.days,0)}</td><td>${fmt(principal)} × ${fmt(x.days,0)} / 30 × ${fmtRate(x.monthlyRate)} / 100</td><td>${formatMoney(x.amount,'TL')}</td></tr>`).join('')}</tbody></table></div><p class="field-help">Kaynak: ${escapeHtml(state.meta.interestSource||'faiz hesaplama Excel’i')} · ${escapeHtml(state.meta.interestMethod||'gün/30 basit faiz')}</p>`;}
function renderDashboard(){if(!$('dashboardKpis'))return;const h=getHistory(),sum=k=>h.reduce((a,x)=>a+toNumber(x[k]),0),totalCost=sum('landedCost'),totalTax=sum('total'),avg=h.length?totalCost/h.length:0;const gtipMap={},originMap={},agreeMap={},taxMap={};for(const x of h){gtipMap[x.gtip]=(gtipMap[x.gtip]||0)+toNumber(x.landedCost);originMap[x.origin]=(originMap[x.origin]||0)+toNumber(x.landedCost);agreeMap[x.agreement]=(agreeMap[x.agreement]||0)+toNumber(x.landedCost);for(const l of x.lines||[])taxMap[l.name]=(taxMap[l.name]||0)+toNumber(l.amount);}$('dashboardKpis').innerHTML=`<div class="data-card"><span>Hesaplama sayısı</span><strong>${h.length}</strong><small>Bu tarayıcı</small></div><div class="data-card"><span>Toplam ithalat maliyeti</span><strong>${formatMoney(totalCost,'TL')}</strong><small>Kayıtlı hesaplamalar</small></div><div class="data-card"><span>Vergi + tamamlama</span><strong>${formatMoney(totalTax,'TL')}</strong><small>Toplam</small></div><div class="data-card"><span>Ortalama maliyet</span><strong>${formatMoney(avg,'TL')}</strong><small>Hesaplama başına</small></div>`;renderBars('taxBreakdown',taxMap);renderBars('topGtips',gtipMap);renderBars('originBreakdown',originMap);renderBars('agreementBreakdown',agreeMap);}
function renderBars(id,map){const el=$(id);if(!el)return;const arr=Object.entries(map).sort((a,b)=>b[1]-a[1]).slice(0,8),max=Math.max(1,...arr.map(x=>x[1]));el.innerHTML=arr.length?arr.map(([k,v])=>`<div class="metric-row"><div><span>${escapeHtml(k)}</span><strong>${formatMoney(v,'TL')}</strong></div><div class="metric-bar"><i style="width:${Math.max(2,v/max*100)}%"></i></div></div>`).join(''):'<div class="empty-history">Rapor için hesaplama geçmişi bulunmuyor.</div>';}
document.addEventListener('DOMContentLoaded',init);
