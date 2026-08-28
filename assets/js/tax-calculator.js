
let TARIFF = {};
let COUNTRIES = {};
let SEARCH_ROWS = [];

const $ = id => document.getElementById(id);
const n = id => Number($(id).value || 0);
const fmt = v => new Intl.NumberFormat('tr-TR',{minimumFractionDigits:2,maximumFractionDigits:2}).format(Number(v||0));
const pct = v => (v === null || v === undefined || v === '') ? '' : `%${Number(v).toLocaleString('tr-TR')}`;

function normTR(s){
  return String(s||'')
    .replace(/\u00A0/g,' ')
    .trim()
    .toLocaleUpperCase('tr-TR');
}
function normGtip(s){
  const digits = String(s||'').replace(/\D/g,'');
  if(digits.length !== 12) return String(s||'').trim();
  return `${digits.slice(0,4)}.${digits.slice(4,6)}.${digits.slice(6,8)}.${digits.slice(8,10)}.${digits.slice(10,12)}`;
}
function num(v){ return (typeof v === 'number' && Number.isFinite(v)) ? v : 0; }
function isPercent(type){ return normTR(type) === 'YÜZDE'; }

async function init(){
  try{
    const [t,u] = await Promise.all([
      fetch('assets/data/tarife-2026.json').then(r=>r.json()),
      fetch('assets/data/ulkeler.json').then(r=>r.json())
    ]);
    TARIFF = t; COUNTRIES = u;
    SEARCH_ROWS = Object.entries(TARIFF).map(([gtip,x])=>({
      gtip, desc:x.d||'', hay:(gtip+' '+(x.d||'')).toLocaleLowerCase('tr-TR')
    }));
    buildCountries();
    bind();
    $('loading').classList.add('hidden');
    $('calculator').classList.remove('hidden');
  }catch(err){
    $('loading').textContent = 'Tarife verileri yüklenemedi. Dosyaları GitHub Pages üzerinde açtığınızdan emin olun.';
    console.error(err);
  }
}

function buildCountries(){
  const origin = $('origin');
  const vals = Object.entries(COUNTRIES)
    .map(([key,x])=>({key, name:x.canonical||key}))
    .sort((a,b)=>a.name.localeCompare(b.name,'tr'));
  origin.innerHTML = vals.map(x=>`<option value="${escapeHtml(x.key)}">${escapeHtml(x.name)}</option>`).join('');
  const germany = vals.find(x=>x.key==='ALMANYA');
  if(germany) origin.value='ALMANYA';
}

function bind(){
  $('gtipSearch').addEventListener('input', onSearch);
  $('gtipSearch').addEventListener('focus', onSearch);
  document.addEventListener('click', e=>{
    if(!e.target.closest('.field')) $('gtipResults').classList.remove('open');
  });
  ['origin','importType','payment','doc','goods','extra','qty','gross','net','dampingManual']
    .forEach(id => $(id).addEventListener('input', calculate));
  $('calcBtn').addEventListener('click', calculate);
  $('resetBtn').addEventListener('click', reset);
}

function onSearch(){
  const raw = $('gtipSearch').value.trim();
  const q = raw.toLocaleLowerCase('tr-TR');
  const box = $('gtipResults');
  if(q.length < 2){ box.classList.remove('open'); box.innerHTML=''; return; }

  const normalized = normGtip(raw);
  let found = SEARCH_ROWS.filter(x => x.gtip.startsWith(normalized) || x.hay.includes(q)).slice(0,10);
  if(!found.length && /\d/.test(raw)){
    const digits = raw.replace(/\D/g,'');
    found = SEARCH_ROWS.filter(x => x.gtip.replace(/\D/g,'').startsWith(digits)).slice(0,10);
  }

  box.innerHTML = found.length
    ? found.map(x=>`<div class="ac-item" data-gtip="${x.gtip}">
        <div class="ac-code">${x.gtip}</div><div class="ac-desc">${escapeHtml(x.desc)}</div>
      </div>`).join('')
    : `<div class="ac-item"><div class="ac-desc">Eşleşme bulunamadı.</div></div>`;
  box.classList.add('open');
  box.querySelectorAll('[data-gtip]').forEach(el=>{
    el.addEventListener('click',()=>{
      selectGtip(el.dataset.gtip);
      box.classList.remove('open');
    });
  });
}

function selectGtip(gtip){
  const x = TARIFF[gtip];
  $('gtip').value = gtip;
  $('gtipSearch').value = gtip;
  $('selectedDesc').textContent = x?.d || '';
  calculate();
}

function originInfo(){
  const key = normTR($('origin').value);
  return COUNTRIES[key] || null;
}

function gtsRate(x, c){
  if(!x) return 0;
  const origin = normTR($('origin').value);
  if(origin === 'BİLİNMEYEN ÜLKE TOPRAKLARI' || origin === 'BILINMEYEN ULKE TOPRAKLARI'){
    const vals = Object.values(x.gts||{}).filter(v=>typeof v==='number');
    return vals.length ? Math.max(...vals) : 0;
  }
  const col = c?.gts;
  return col && typeof x.gts?.[col] === 'number' ? x.gts[col] : 0;
}

function calculate(){
  const gtip = $('gtip').value;
  const x = TARIFF[gtip];
  const c = originInfo();

  const goods = n('goods');
  const extra = n('extra');
  const qty = n('qty');
  const gross = n('gross');
  const damping = n('dampingManual');

  let gozDiff = 0;
  if(x && num(x.goz_r) > 0){
    const unit = normTR(x.goz_u);
    const base = unit === 'ADET' ? qty : gross;
    gozDiff = Math.max(0, num(x.goz_r) * base - (goods + extra));
  }
  const cif = goods + extra + gozDiff;
  $('gozetimDiff').value = fmt(gozDiff);
  $('cif').value = fmt(cif);

  if(!x){
    clearResults();
    $('resultStatus').className='status warn';
    $('resultStatus').textContent='GTİP BEKLENİYOR';
    return;
  }

  const importType = normTR($('importType').value);
  const group = normTR(c?.group);
  const origin = normTR($('origin').value);

  let gv = 0;
  if(importType === 'DI'){
    if(isPercent(x.gv_t)) gv = cif * num(x.gv) / 100;
  }

  let igv = 0;
  if(group !== 'AB' && group !== 'TÜRKİYE'){
    if(importType === 'DI'){
      if(isPercent(x.igv_t)) igv = cif * num(x.igv) / 100;
    }else if(importType === 'AT' || importType === 'A.TR'){
      if(normTR(c?.igv_atr) === 'EVET' && isPercent(x.igv_t)){
        igv = cif * num(x.igv) / 100;
      }
    }
  }

  const emyRate = gtsRate(x,c);
  let emy = 0;
  if((importType === 'AT' || importType === 'A.TR') && gv === 0 && isPercent(x.emy_t)){
    emy = cif * emyRate / 100;
  }

  const isUS = ['BİRLEŞİK DEVLETLER','AMERİKA BİRLEŞİK DEVLETLERİ','ABD','A.B.D.'].includes(origin)
    || normTR(c?.canonical) === 'BİRLEŞİK DEVLETLER';
  const abd = isUS ? cif * num(x.abd) / 100 : 0;

  const band = cif * num(x.band);

  let otv = 0;
  if(isPercent(x.otv_t)){
    otv = (cif + gv + igv + emy + abd + band) * num(x.otv) / 100;
  }

  const payment = normTR($('payment').value);
  const kkdfApplies = ['MAL MUKABİLİ','KABUL KREDİLİ','VADELİ AKREDİTİF'].includes(payment);
  const kkdf = kkdfApplies ? cif * num(x.kkdf) / 100 : 0;

  const kdvBase = cif + gv + igv + emy + abd + band + otv + kkdf + damping;
  const kdv = kdvBase * num(x.kdv) / 100;

  const total = gv + igv + emy + abd + band + otv + kkdf + kdv + damping;

  setResult('gv',gv); setResult('igv',igv); setResult('emy',emy); setResult('abd',abd);
  setResult('band',band); setResult('otv',otv); setResult('kkdf',kkdf); setResult('kdv',kdv);
  setResult('damping',damping);
  $('totalTax').textContent=fmt(total);

  $('gvRate').textContent=pct(x.gv);
  $('igvRate').textContent=pct(x.igv);
  $('emyRate').textContent=pct(emyRate);
  $('abdRate').textContent=pct(x.abd);
  $('bandRate').textContent=x.band ? `× ${Number(x.band).toLocaleString('tr-TR')}` : '';
  $('otvRate').textContent=pct(x.otv);
  $('kkdfRate').textContent=pct(x.kkdf);
  $('kdvRate').textContent=pct(x.kdv);

  $('gtipMeta').textContent = `${gtip} • ${x.d || 'Tanım yok'}`;
  $('originMeta').textContent = c
    ? `Menşe grubu: ${c.group || '—'}${c.doc ? ` • Önerilen belge: ${c.doc}` : ''}`
    : 'Menşe ülke listesinde bulunamadı.';

  const warnings = buildWarnings(x,c,importType,damping);
  $('warnings').innerHTML = warnings.length
    ? warnings.map(w=>`<li>${escapeHtml(w)}</li>`).join('')
    : '<li>Otomatik kontrollerde ek uyarı oluşmadı.</li>';

  const ok = warnings.length === 0;
  $('resultStatus').className='status ' + (ok?'good':'warn');
  $('resultStatus').textContent=ok?'OTOMATİK KONTROL OK':'MANUEL KONTROL';

  $('ruleText').textContent =
    `${(importType==='AT'||importType==='A.TR')?'A.TR/AT':'Dİ'} | MENŞE GRUBU: ${c?.group||'BULUNAMADI'} | `+
    `GV=${fmt(gv)} | İGV=${fmt(igv)} | EMY-GTS=${fmt(emy)} | KDV=${fmt(kdv)}`;
}

function buildWarnings(x,c,importType,damping){
  const w=[];
  if(!c) w.push('MENŞE LİSTEDE YOK.');
  if(x.gv === null || x.gv === undefined) w.push('GTİP / GV ORANI BULUNAMADI.');
  if(normTR(x.manuel)==='EVET') w.push('GTİP ÖZEL KURAL İÇERİYOR.');
  if(x.gv_t && !isPercent(x.gv_t)) w.push('GV MAKTU/KARMA; MANUEL KONTROL.');
  if(x.igv_t && !isPercent(x.igv_t)) w.push('İGV MAKTU/KARMA; MANUEL KONTROL.');
  if(x.otv_t && !isPercent(x.otv_t)) w.push('ÖTV MAKTU/KARMA; MANUEL KONTROL.');
  if(gtsRate(x,c)!==0 && x.emy_t && !isPercent(x.emy_t)) w.push('EMY-GTS MAKTU/KARMA; MANUEL KONTROL.');
  if(x.damp) w.push('DAMPİNG NOTU BULUNUYOR; DAMPİNG TUTARINI MANUEL GİRİN.');
  if(x.ozel) w.push('ÖZEL MEVZUAT / KURAL KONTROLÜ GEREKİYOR.');
  if(normTR($('doc').value)!=='YOK') w.push('TERCİHLİ MENŞE BELGESİ ORANI MANUEL DOĞRULANMALI.');
  if(!['DI','AT','A.TR'].includes(importType)) w.push('İTHALAT TÜRÜ GEÇERSİZ.');
  if((importType==='AT'||importType==='A.TR') && !['AB','TÜRKİYE'].includes(normTR(c?.group))
     && normTR(c?.igv_atr)!=='EVET' && num(x.igv)>0){
    w.push('İGV A.TR ÜLKE KURALI MANUEL KONTROL EDİLMELİ.');
  }
  if(damping>0 && !x.damp) w.push('DAMPİNG MANUEL TUTAR GİRİLDİ; DAYANAK KONTROL EDİLMELİ.');
  return w;
}

function setResult(id,v){ $(id).textContent=fmt(v); }

function clearResults(){
  ['gv','igv','emy','abd','band','otv','kkdf','kdv','damping'].forEach(id=>setResult(id,0));
  ['gvRate','igvRate','emyRate','abdRate','bandRate','otvRate','kkdfRate','kdvRate'].forEach(id=>$(id).textContent='');
  $('totalTax').textContent='0,00';
  $('gtipMeta').textContent='Henüz GTİP seçilmedi.';
  $('originMeta').textContent='';
  $('warnings').innerHTML='<li>Hesaplama için GTİP seçin.</li>';
  $('ruleText').textContent='—';
}

function reset(){
  $('gtip').value='';
  $('gtipSearch').value='';
  $('selectedDesc').textContent='GTİP seçilmedi.';
  ['goods','extra','qty','gross','net','dampingManual'].forEach(id=>$(id).value=0);
  $('payment').value='PEŞİN';
  $('importType').value='AT';
  $('doc').value='YOK';
  $('gozetimDiff').value='0,00';
  $('cif').value='0,00';
  clearResults();
  $('resultStatus').className='status warn';
  $('resultStatus').textContent='GTİP BEKLENİYOR';
}

function escapeHtml(s){
  return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]));
}
init();
