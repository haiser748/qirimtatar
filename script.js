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
    if (q) arr = arr.filter(x => norm(x.word).includes(q) || norm(x.translation).includes(q) || norm(x.transcription).includes(q)||norm(x.example).includes(q)||norm(x.category).includes(q));
  const s=$("sortSelect").value;if(s==="az")arr.sort((a,b)=>String(a.word).localeCompare(String(b.word),"ar"));if(s==="za")arr.sort((a,b)=>String(b.word).localeCompare(String(a.word),"ar"));
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
function norm(v){return String(v??"").trim().toLowerCase()}
function esc(v){return String(v??"").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;")}
const savedTheme=localStorage.getItem("dictionaryTheme");if(savedTheme)document.documentElement.dataset.theme=savedTheme;
$("themeBtn").onclick=()=>{const dark=document.documentElement.dataset.theme==="dark";document.documentElement.dataset.theme=dark?"":"dark";localStorage.setItem("dictionaryTheme",dark?"":"dark");$("themeBtn").textContent=dark?"☾":"☀"};
$("themeBtn").textContent=document.documentElement.dataset.theme==="dark"?"☀":"☾";
loadDictionary();
