let dictionary=[], favorites=new Set(), currentView="all", currentCategory="";
const $=id=>document.getElementById(id);
const input=$("searchInput"), results=$("results"), status=$("status"), empty=$("empty");

try{favorites=new Set(JSON.parse(localStorage.getItem("dictionaryFavorites")||"[]"))}catch{}

async function loadDictionary(){
  try{
      const response = await fetch("dictionary.xlsx?v=" + Date.now(), { cache: "no-store" });
    if(!response.ok)throw new Error("Excel не найден");
    const buffer=await response.arrayBuffer();
    const wb=XLSX.read(buffer,{type:"array"});
    const sheet=wb.Sheets[wb.SheetNames[0]];
    dictionary=XLSX.utils.sheet_to_json(sheet,{defval:""}).map((x,i)=>({...x,_id:String(i)}));
    console.log(dictionary.length);
    buildCategories(); updateCounts(); render();
  }catch(e){
    status.textContent="Не удалось загрузить слова.";
    console.error(e);
  }
}
function buildCategories(){
  const map={}; dictionary.forEach(x=>{const c=String(x.category||"Без категории");map[c]=(map[c]||0)+1});
  $("categories").innerHTML=Object.entries(map).sort().map(([c,n])=>
    `<button class="nav-item" data-category="${esc(c)}">• <span>${esc(c)}</span><b>${n}</b></button>`).join("");
  $("categories").querySelectorAll("[data-category]").forEach(b=>b.onclick=()=>{currentView="category";currentCategory=b.dataset.category;activateCategory(b);render()});
}
function updateCounts(){$("countAll").textContent=dictionary.length;$("countFav").textContent=favorites.size}
function activateCategory(btn){document.querySelectorAll(".nav-item").forEach(x=>x.classList.remove("active"));btn.classList.add("active")}
document.querySelectorAll("[data-view]").forEach(b=>b.onclick=()=>{currentView=b.dataset.view;currentCategory="";document.querySelectorAll(".nav-item").forEach(x=>x.classList.remove("active"));b.classList.add("active");render()});
$("sortSelect").onchange=render;
input.oninput=()=>{ $("clearBtn").hidden=!input.value; render() };
$("clearBtn").onclick=()=>{input.value="";$("clearBtn").hidden=true;render();input.focus()};
document.querySelectorAll(".quick button").forEach(b=>b.onclick=()=>{input.value=b.dataset.query;$("clearBtn").hidden=false;render();window.scrollTo({top:0,behavior:"smooth"})});
document.addEventListener("keydown",e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="k"){e.preventDefault();input.focus()}});
function render(){
  let arr=[...dictionary], q=norm(input.value);
  if(currentView==="favorites")arr=arr.filter(x=>favorites.has(x._id));
  if(currentView==="category")arr=arr.filter(x=>String(x.category||"Без категории")===currentCategory);
  if(q){
    // Сначала самые подходящие результаты.
    arr = arr
      .map(x => ({item:x, score:searchScore(q,x)}))
      .filter(x => x.score > 0)
      .sort((a,b) => b.score - a.score)
      .map(x => x.item);
  }
  const s=$("sortSelect").value;
  // При поиске сначала показываем релевантность, а не алфавит.
  if(!q){
    if(s==="az")arr.sort((a,b)=>String(a.word).localeCompare(String(b.word),"ar"));
    if(s==="za")arr.sort((a,b)=>String(b.word).localeCompare(String(a.word),"ar"));
  }
  $("viewTitle").textContent=q?`Поиск: «${input.value}»`:currentView==="favorites"?"Избранное":currentView==="category"?currentCategory:"Все слова";
  status.textContent=`Найдено: ${arr.length}`;
  results.innerHTML=arr.slice(0,100).map(card).join("");
  empty.hidden=arr.length!==0;results.hidden=arr.length===0;
  results.querySelectorAll(".favorite").forEach(b=>b.onclick=()=>toggleFav(b.dataset.id));
}
function card(x){
  const on=favorites.has(x._id);
  return `<article class="card">
    <div class="card-top"><div><div class="word">${esc(x.word)}</div><div class="translation">${esc(x.translation)}</div>${x.category?`<span class="category">${esc(x.category)}</span>`:""}</div>
    <button class="favorite ${on?"on":""}" data-id="${x._id}" title="Избранное">${on?"♥":"♡"}</button></div>
    ${x.transcription ? `<div class="line"></div><div class="label">Транскрипция</div><div class="transcription">${esc(x.transcription)}</div>`:""}
    ${x.example?`<div class="line"></div><div class="label">Пример</div><div class="example">${esc(x.example)}</div>`:""}
  </article>`;
}
function toggleFav(id){favorites.has(id)?favorites.delete(id):favorites.add(id);localStorage.setItem("dictionaryFavorites",JSON.stringify([...favorites]));updateCounts();render()}

// Нечёткий поиск: находит слова даже при небольших опечатках,
// переставленных/пропущенных символах и неполном вводе.
// Оценка релевантности результата. Чем выше балл — тем ближе результат к запросу.
function searchScore(query, x){
  const fields=[
    {value:x.word, weight:100},
    {value:x.translation, weight:75},
    {value:x.transcription, weight:65},
    {value:x.example, weight:40},
    {value:x.category, weight:20}
  ];
  let best=0;
  for(const {value,weight} of fields){
    const field=norm(value);
    if(!field) continue;
    const score=fuzzyFieldScore(query,field);
    best=Math.max(best, score*weight/100);
  }
  return best;
}

function fuzzyFieldScore(query, field){
  if(!query || !field) return 0;
  if(field===query) return 1.00;
  if(field.startsWith(query)) return 0.95;
  if(field.includes(query)) return 0.88;

  const words=field.split(/\s+/).filter(Boolean);
  let best=0;
  for(const word of words){
    if(word===query) best=Math.max(best,1);
    else if(word.startsWith(query)) best=Math.max(best,0.94);
    else if(word.includes(query)) best=Math.max(best,0.86);
    else best=Math.max(best, fuzzyDistanceScore(query,word));
  }

  // Для фразы с несколькими словами учитываем совпадение всех частей.
  const parts=query.split(/\s+/).filter(Boolean);
  if(parts.length>1){
    const scores=parts.map(part=>Math.max(...words.map(word=>fuzzyDistanceScore(part,word)),0));
    if(scores.length && scores.every(v=>v>0)) best=Math.max(best, scores.reduce((a,b)=>a+b,0)/scores.length);
  }
  return best;
}

function fuzzyDistanceScore(query,target){
  if(target.includes(query)) return 0.86;
  const maxDistance=query.length<=3 ? 1 : query.length<=6 ? 2 : 3;
  const distance=levenshtein(query,target);
  if(distance<=maxDistance){
    return Math.max(0.35, 0.84 - distance*0.13 - Math.abs(target.length-query.length)*0.015);
  }
  if(query.length>=4){
    const len=query.length;
    let best=0;
    for(let i=0;i<=target.length-len;i++){
      const d=levenshtein(query,target.slice(i,i+len));
      if(d<=Math.min(2,maxDistance)) best=Math.max(best,0.72-d*0.13);
    }
    return best;
  }
  return 0;
}

function exactMatch(query, x){
  return [x.word,x.translation,x.transcription,x.example,x.category].map(norm).some(field => field === query);
}

function fuzzyScore(query, x){
  const fields=[x.word,x.translation,x.transcription,x.example,x.category].map(norm).filter(Boolean);
  let best=0;
  for(const field of fields){
    if(field===query) return 10000;
    if(field.startsWith(query)) best=Math.max(best,9000-query.length);
    else if(field.includes(query)) best=Math.max(best,8000-query.length);
    for(const token of field.split(/\s+/)){
      if(!token) continue;
      const d=levenshtein(query,token);
      const max=query.length<=3?1:query.length<=6?2:3;
      if(d<=max) best=Math.max(best,7000-d*500-Math.abs(token.length-query.length)*20);
    }
  }
  return best;
}

function fuzzyMatch(query, x){
  const fields=[x.word,x.translation,x.transcription,x.example,x.category].map(norm).filter(Boolean);
  return fields.some(field=>fuzzyFieldMatch(query,field));
}

function fuzzyFieldMatch(query, field){
  if(!query || !field) return false;
  if(field.includes(query)) return true;

  // Проверяем отдельные слова, чтобы опечатка в одном слове не мешала поиску.
  const words=field.split(/\s+/).filter(Boolean);
  const parts=query.split(/\s+/).filter(Boolean);
  if(parts.length>1 && parts.every(part=>words.some(word=>fuzzyTokenMatch(part,word)))) return true;
  return fuzzyTokenMatch(query,field);
}

function fuzzyTokenMatch(query, target){
  if(target.includes(query)) return true;
  const maxDistance=query.length<=3 ? 1 : query.length<=6 ? 2 : 3;
  const distance=levenshtein(query,target);
  if(distance<=maxDistance) return true;

  // Для длинных слов разрешаем искать близкий фрагмент внутри слова.
  if(query.length>=4){
    const len=query.length;
    for(let i=0;i<=target.length-len;i++){
      if(levenshtein(query,target.slice(i,i+len))<=Math.min(2,maxDistance)) return true;
    }
  }
  return false;
}

function levenshtein(a,b){
  if(a===b) return 0;
  if(!a.length) return b.length;
  if(!b.length) return a.length;
  if(a.length>b.length){const t=a;a=b;b=t;}
  let prev=Array.from({length:a.length+1},(_,i)=>i);
  for(let j=1;j<=b.length;j++){
    const cur=[j];
    for(let i=1;i<=a.length;i++){
      const cost=a[i-1]===b[j-1]?0:1;
      cur[i]=Math.min(cur[i-1]+1,prev[i]+1,prev[i-1]+cost);
    }
    prev=cur;
  }
  return prev[a.length];
}

function norm(v){return String(v??"").trim().toLowerCase()}
function esc(v){return String(v??"").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;")}
const savedTheme=localStorage.getItem("dictionaryTheme");if(savedTheme)document.documentElement.dataset.theme=savedTheme;
$("themeBtn").onclick=()=>{const dark=document.documentElement.dataset.theme==="dark";document.documentElement.dataset.theme=dark?"":"dark";localStorage.setItem("dictionaryTheme",dark?"":"dark");$("themeBtn").textContent=dark?"☾":"☀"};
$("themeBtn").textContent=document.documentElement.dataset.theme==="dark"?"☀":"☾";
loadDictionary();
