const PASSPORT_KEY = 'fruitlab-explorer-passport-v1';
const STAR_KEY = 'fruitlab-explorer-stars-v1';
const LAST_FRUIT_KEY = 'fruitlab-last-fruit-v1';
const names = {apple:'Măr',banana:'Banană',orange:'Portocală',pear:'Pară',peach:'Piersică',mango:'Mango',pineapple:'Ananas',strawberry:'Căpșună',grape:'Struguri',watermelon:'Pepene verde',lemon:'Lămâie',lime:'Lime',kiwi:'Kiwi',plum:'Prună',cherry:'Cireașă',coconut:'Nucă de cocos',papaya:'Papaya',pomegranate:'Rodie',avocado:'Avocado',apricot:'Caisă',blueberry:'Afină',raspberry:'Zmeură',blackberry:'Mură',grapefruit:'Grepfrut',fig:'Smochină',guava:'Guava','dragon fruit':'Fructul dragonului','passion fruit':'Fructul pasiunii',lychee:'Litchi',nectarine:'Nectarină',cranberry:'Merișor',tangerine:'Mandarină',clementine:'Clementină',persimmon:'Curmal japonez',date:'Curmală',durian:'Durian',jackfruit:'Jackfruit',starfruit:'Carambolă',longan:'Longan',rambutan:'Rambutan',quince:'Gutui',mulberry:'Dudă',gooseberry:'Agrișă',pomelo:'Pomelo',kumquat:'Kumquat',tamarind:'Tamarind',cantaloupe:'Pepene galben',honeydew:'Pepene honeydew',plantain:'Banană plantain',soursop:'Guanabana','custard apple':'Anona',elderberry:'Soc',mangosteen:'Mangostan',blackcurrant:'Coacăză neagră',redcurrant:'Coacăză roșie',greengage:'Reine Claude','blood orange':'Portocală roșie','prickly pear':'Smochin de India',cherimoya:'Cherimoya',medlar:'Moșmon',olive:'Măslină'};
const englishNames = {grape:'Grapes',strawberry:'Strawberry',blueberry:'Blueberry',raspberry:'Raspberry',blackberry:'Blackberry',cherry:'Cherry',kiwi:'Kiwi fruit','dragon fruit':'Dragon fruit','passion fruit':'Passion fruit',starfruit:'Star fruit','custard apple':'Custard apple',date:'Date',quince:'Quince',mulberry:'Mulberry',gooseberry:'Gooseberry',cantaloupe:'Cantaloupe',honeydew:'Honeydew melon','blood orange':'Blood orange','prickly pear':'Prickly pear',blackcurrant:'Blackcurrant',redcurrant:'Redcurrant',greengage:'Greengage'};
const emoji = {apple:'🍎',banana:'🍌',orange:'🍊',pear:'🍐',peach:'🍑',mango:'🥭',pineapple:'🍍',strawberry:'🍓',grape:'🍇',watermelon:'🍉',lemon:'🍋',kiwi:'🥝',cherry:'🍒',coconut:'🥥',avocado:'🥑',blueberry:'🫐','dragon fruit':'🐉'};
const facts = {apple:'Mărul face parte din familia trandafirilor.',banana:'Botanic, banana este o bacă.',strawberry:'Semințele mici de la exterior sunt fructele propriu-zise.',pineapple:'Ananasul este alcătuit din multe fructe mici unite.',watermelon:'Pepenele verde este înrudit cu castravetele.',avocado:'Avocado este un fruct, chiar dacă apare des în mâncăruri sărate.',kiwi:'Miezul verde ascunde semințe mici, comestibile.',orange:'Portocala este un fruct citric, înrudit cu lămâia.'};
const read=(key,fallback)=>{try{return JSON.parse(localStorage.getItem(key))??fallback;}catch{return fallback;}};
const foundFruits=()=>{const value=read(PASSPORT_KEY,[]);return Array.isArray(value)?value:[];};
const getStars=()=>Number(read(STAR_KEY,0))||0;
const en=id=>englishNames[id]||id.replace(/\b\w/g,c=>c.toUpperCase());

// Shared navigation and progress badge on every page.
const menu=document.querySelector('#menu-toggle'),header=document.querySelector('.topbar');
menu?.addEventListener('click',()=>{const open=header.classList.toggle('menu-open');menu.setAttribute('aria-expanded',String(open));menu.setAttribute('aria-label',open?'Închide meniul':'Deschide meniul');});
document.querySelectorAll('.main-nav a').forEach(link=>{
  if(link.dataset.nav===document.body.dataset.page){link.classList.add('route-active');link.setAttribute('aria-current','page');}
  link.addEventListener('click',()=>{header.classList.remove('menu-open');menu?.setAttribute('aria-expanded','false');});
});
document.querySelectorAll('.passport-count').forEach(node=>node.textContent=foundFruits().length);
document.querySelectorAll('.star-count,.passport-stars').forEach(node=>node.textContent=getStars());

const surpriseButton=document.querySelector('#fruit-surprise');
if(surpriseButton){
  const surprises=[
    ['🍍','Știai că ananasul e format din multe fructe?','Bobițele lui s-au unit într-un fruct mare și aromat.'],
    ['🍓','Căpșuna își poartă semințele la exterior!','Punctele mici de pe coajă sunt fructe minuscule.'],
    ['🍌','Banana este o bacă adevărată!','Din punct de vedere botanic, banana este o bacă, iar căpșuna nu.'],
    ['🍊','Portocala face parte din familia citricelor.','Lămâia și mandarina sunt verișoarele ei parfumate.'],
    ['🥝','Kiwi are semințe comestibile.','Miezul verde ascunde multe semințe mici, negre și crocante.'],
    ['🍉','Pepenele verde are foarte multă apă.','De aceea este atât de răcoritor într-o zi călduroasă!']
  ];let previous=-1;
  surpriseButton.addEventListener('click',()=>{let index;do{index=Math.floor(Math.random()*surprises.length);}while(index===previous);previous=index;const [,title,fact]=surprises[index];document.querySelector('#day-fruit-emoji').textContent=String(index+1).padStart(2,'0');document.querySelector('#day-fruit-title').textContent=title;document.querySelector('#day-fruit-fact').textContent=fact;surpriseButton.classList.remove('is-spinning');void surpriseButton.offsetWidth;surpriseButton.classList.add('is-spinning');});
}

// Fruitopedia cards open a playful fact card with a path back to the scanner.
document.querySelectorAll('.filter').forEach(button=>button.addEventListener('click',()=>{
  const category=button.dataset.filter;let visible=0;
  document.querySelectorAll('.filter').forEach(item=>{const active=item===button;item.classList.toggle('active',active);item.setAttribute('aria-pressed',String(active));});
  document.querySelectorAll('.fruit-card').forEach(card=>{const show=category==='toate'||card.dataset.category.split(' ').includes(category);card.hidden=!show;if(show)visible++;});
  const empty=document.querySelector('.no-results');if(empty)empty.hidden=visible>0;
}));
const fruitDialog=document.querySelector('#fruit-dialog');
function openFruitCard(card){
  const id=card.dataset.fruitId;if(!id||!names[id])return;
  const panel=document.querySelector('#fruit-dialog-content');
  const image=card.querySelector('img')?.src||'';panel.innerHTML=`<img class="dialog-fruit-image" src="${image}" alt="${names[id]}"><p class="dialog-kicker">FIȘĂ DE EXPLORATOR</p><h2>${names[id]} <small>${en(id)}</small></h2><p>${facts[id]||`${names[id]} este unul dintre fructele pe care le poți descoperi cu FruitLab AR.`}</p><a class="primary-link" href="scan.html?challenge=${encodeURIComponent(id)}">Găsește-l cu camera <span>↗</span></a>`;
  fruitDialog?.showModal();
}
document.querySelectorAll('.fruit-card[data-fruit-id]').forEach(card=>{
  card.tabIndex=0;card.setAttribute('role','button');card.setAttribute('aria-label',`Află o curiozitate despre ${names[card.dataset.fruitId]}`);
  card.addEventListener('click',()=>openFruitCard(card));card.addEventListener('keydown',event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();openFruitCard(card);}});
});
document.querySelector('#dialog-close')?.addEventListener('click',()=>fruitDialog.close());
fruitDialog?.addEventListener('click',event=>{if(event.target===fruitDialog)fruitDialog.close();});

// Pașaportul records new fruit discoveries and explorer levels.
const passportGrid=document.querySelector('#passport-grid');
if(passportGrid){
  const found=foundFruits(),count=found.length,total=document.querySelector('#passport-total');
  if(total)total.textContent=`${count} ${count===1?'fruct descoperit':'fructe descoperite'}`;
  const message=document.querySelector('#passport-message');if(message)message.textContent=count?`Ai completat ${count} descoperiri. Continuă aventura!`:'Pornește prima aventură ca să completezi pașaportul!';
  const levels=[3,6,12],next=levels.find(level=>count<level)||15,previous=levels[levels.indexOf(next)-1]||0,progress=Math.max(0,Math.min(100,((count-previous)/(next-previous))*100));
  const level=document.querySelector('#passport-level'),label=document.querySelector('#passport-progress-label'),bar=document.querySelector('#passport-progress-bar');
  if(level)level.textContent=count>=12?'Nivel 4 · Campion FruitLab':count>=6?'Nivel 3 · Explorator de soiuri':count>=3?'Nivel 2 · Culegător curios':'Nivel 1 · Explorator începător';
  if(label)label.textContent=`${Math.max(0,next-count)} fructe până la următoarea insignă`;
  if(bar)bar.style.width=`${progress}%`;
  document.querySelectorAll('.passport-badges article').forEach(card=>{const unlocked=count>=Number(card.dataset.badge);card.classList.toggle('is-unlocked',unlocked);card.setAttribute('aria-label',`${card.querySelector('strong').textContent}: ${unlocked?'deblocată':'încă blocată'}`);});
  passportGrid.replaceChildren();
  if(!count){const empty=document.createElement('div');empty.className='passport-empty';empty.innerHTML='<strong>Jurnalul tău începe aici</strong><small>Scanează primul fruct pentru a începe colecția.</small>';passportGrid.append(empty);}
  found.forEach(id=>{if(!names[id])return;const card=document.createElement('article');card.className='discovery-record';card.innerHTML=`<span class="record-index">${String(passportGrid.querySelectorAll('.discovery-record').length+1).padStart(2,'0')}</span><div><strong>${names[id]}</strong><small>${en(id)}</small></div><i>DESCOPERIT</i>`;passportGrid.append(card);});
}

// Short quiz rounds: scanned-fruit translation, then rotating fruit facts.
if(document.body.dataset.page==='play'){
  const question=document.querySelector('#quiz-question'),options=document.querySelector('#quiz-options'),feedback=document.querySelector('#quiz-feedback'),status=document.querySelector('#quiz-status'),nextButton=document.querySelector('#next-question'),startButton=document.querySelector('#start-surprise');
  let round=1,fruitId=new URLSearchParams(location.search).get('fruit')||localStorage.getItem(LAST_FRUIT_KEY),answered=false;
  const allIds=Object.keys(names),randomId=()=>allIds[Math.floor(Math.random()*allIds.length)];
  function setRound(id,number=round){
    fruitId=id||randomId();round=number;answered=false;options.replaceChildren();nextButton.hidden=true;startButton.hidden=true;
    document.querySelector('#quiz-round-label').textContent=`Runda ${round}`;document.querySelector('#quiz-progress-bar').style.width=`${((round-1)%5+1)*20}%`;
    document.querySelector('#quiz-mascot').textContent=['FL','AI','AR','RO','EN'][(round-1)%5];
    let correct=id,questionText,englishMode=false;
    const mode=(round-1)%5;
    if(mode===0){questionText=`Cum se spune „${names[id]}” în limba engleză?`;englishMode=true;}
    else if(mode===1){questionText='Care dintre fructe este o bacă din punct de vedere botanic?';correct='banana';}
    else if(mode===2){questionText='La ce fruct găsim semințele mici la exterior?';correct='strawberry';}
    else if(mode===3){questionText='Care dintre acestea este un fruct citric?';correct='orange';}
    else{questionText='Ce fruct crește în ciorchini?';correct='grape';}
    question.textContent=questionText;status.textContent=`MISIUNE ${String(mode+1).padStart(2,'0')} · FRUCTE` ;feedback.textContent='Alege răspunsul. Fiecare încercare te învață ceva nou!';
    const choices=[correct];while(choices.length<4){const choice=randomId();if(!choices.includes(choice))choices.push(choice);}choices.sort(()=>Math.random()-.5);
    choices.forEach(choice=>{const button=document.createElement('button');button.className='quiz-option';button.type='button';button.textContent=englishMode?en(choice):names[choice];button.addEventListener('click',()=>{
      if(answered)return;answered=true;options.querySelectorAll('button').forEach(item=>item.disabled=true);
      if(choice===correct){button.classList.add('correct');feedback.textContent='Bravo! Răspuns corect. Ai câștigat un punct!';localStorage.setItem(STAR_KEY,String(getStars()+1));document.querySelectorAll('.star-count,.passport-stars').forEach(node=>node.textContent=getStars());}
      else{button.classList.add('incorrect');options.querySelectorAll('button').forEach(item=>{const expected=englishMode?en(correct):names[correct];if(item.textContent===expected)item.classList.add('correct');});feedback.textContent=`Aproape! Răspunsul era ${englishMode?en(correct):names[correct]}. Acum îl știi! 🌟`;}
      nextButton.hidden=false;
    });options.append(button);});
  }
  startButton?.addEventListener('click',()=>setRound(randomId(),1));
  nextButton?.addEventListener('click',()=>setRound(randomId(),round+1));
  if(fruitId&&names[fruitId])setRound(fruitId,1);else startButton.hidden=false;
}
