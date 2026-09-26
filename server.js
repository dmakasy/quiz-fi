// Quiz "tout sur le telephone" - chaque joueur joue sur son appareil.
// Serveur Node.js + Socket.IO. Classement commun en temps reel.
const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const PORT = process.env.PORT || 3000;
const app = express();
const server = http.createServer(app);
const io = new Server(server);
app.use(express.static(path.join(__dirname, 'public')));

let QUESTIONS = [];
let QUIZ_TITLE = 'Quiz';

/* ---------- Analyse tolerante ---------- */
function norm(s){ return String(s==null?'':s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,''); }
function letterToIndex(v){ const s=String(v==null?'':v).trim().toUpperCase(); if(s.length===1 && 'ABCDE'.includes(s)) return 'ABCDE'.indexOf(s); const n=parseInt(s,10); if(n>=1&&n<=5) return n-1; return -1; }
function normImg(img){ img=String(img==null?'':img).trim(); if(!img) return ''; if(/^https?:/i.test(img)) return img; if(img.indexOf('/')>=0) return '/'+img.replace(/^\/+/,''); return '/images/'+img; }
function rowsToQuestions(rows){
  const out=[];
  rows.forEach(r=>{
    const m={}; Object.keys(r).forEach(k=>m[norm(k)]=r[k]);
    const q=m['question']||m['questiontext']||m['enonce']||m['q']||'';
    const o=[
      m['optiona']||m['answer1']||m['option1']||m['a']||m['reponse1']||m['choix1']||'',
      m['optionb']||m['answer2']||m['option2']||m['b']||m['reponse2']||m['choix2']||'',
      m['optionc']||m['answer3']||m['option3']||m['c']||m['reponse3']||m['choix3']||'',
      m['optiond']||m['answer4']||m['option4']||m['d']||m['reponse4']||m['choix4']||''
    ].map(x=>String(x==null?'':x).trim());
    const corr=m['bonnereponse']||m['correctanswers']||m['correctanswer']||m['correct']||m['reponsecorrecte']||m['bonne']||m['solution']||m['reponsen']||'';
    let a=letterToIndex(corr);
    if(a<0 && corr){ const ci=o.findIndex(x=>x && norm(x)===norm(corr)); if(ci>=0)a=ci; }
    const img=normImg(m['image']||m['img']||m['imagelink']||m['imagefacultatif']||'');
    const opts=o.filter(x=>x!=='');
    if(q && opts.length>=2 && a>=0 && a<opts.length) out.push({ q:String(q).trim(), o:opts, a, img });
  });
  return out;
}
function workbookToQuiz(wb){ const ws=wb.Sheets[wb.SheetNames[0]]; return rowsToQuestions(XLSX.utils.sheet_to_json(ws,{defval:''})); }
function loadFromDisk(){
  try{
    const xp=path.join(__dirname,'questions.xlsx'), cp=path.join(__dirname,'questions.csv'), jp=path.join(__dirname,'questions.json');
    if(fs.existsSync(xp)) return { title:'Quiz', questions:workbookToQuiz(XLSX.readFile(xp)) };
    if(fs.existsSync(cp)) return { title:'Quiz', questions:workbookToQuiz(XLSX.read(fs.readFileSync(cp,'utf8'),{type:'string'})) };
    if(fs.existsSync(jp)){
      const d=JSON.parse(fs.readFileSync(jp,'utf8'));
      const questions=(d.questions||d.q||[]).map(x=>({q:x.q||x.question,o:x.o||x.options,a:(x.a!==undefined?x.a:x.correct),img:normImg(x.img||x.image||'')})).filter(v=>v.q&&Array.isArray(v.o)&&v.o.length>=2&&typeof v.a==='number');
      return { title:d.title||d.t||'Quiz', questions };
    }
  }catch(e){ console.error('Chargement questions :', e.message); }
  return { title:'Quiz', questions:[] };
}
const https=require('https');
function fetchText(url, redirects){
  redirects=redirects||0;
  return new Promise((resolve,reject)=>{
    if(redirects>5) return reject(new Error('Trop de redirections'));
    let lib; try{ lib=url.startsWith('https')?https:http; }catch(_){ return reject(new Error('URL invalide')); }
    lib.get(url,{headers:{'User-Agent':'quiz-fi'}},res=>{
      if(res.statusCode>=300 && res.statusCode<400 && res.headers.location){ res.resume(); return resolve(fetchText(new URL(res.headers.location,url).toString(), redirects+1)); }
      if(res.statusCode!==200){ res.resume(); return reject(new Error('HTTP '+res.statusCode)); }
      let data=''; res.setEncoding('utf8'); res.on('data',c=>data+=c); res.on('end',()=>resolve(data));
    }).on('error',reject);
  });
}
let SOURCE='disk';
async function loadQuestions(){
  const url=(process.env.QUESTIONS_URL||'').trim();
  if(url){
    try{
      const text=await fetchText(url);
      const qs=workbookToQuiz(XLSX.read(text,{type:'string'}));
      if(qs.length){ QUESTIONS=qs; QUIZ_TITLE=process.env.QUIZ_TITLE||QUIZ_TITLE||'Quiz'; SOURCE='url'; return {source:'url',count:qs.length}; }
      console.error('QUESTIONS_URL : 0 question detectee');
    }catch(e){ console.error('QUESTIONS_URL echec :', e.message); }
  }
  const r=loadFromDisk(); QUESTIONS=r.questions; QUIZ_TITLE=r.title; SOURCE='disk'; return {source:'disk',count:r.questions.length};
}

/* ---------- Pages ---------- */
app.get('/', (req,res)=>res.sendFile(path.join(__dirname,'public','play.html')));
app.get('/board', (req,res)=>res.sendFile(path.join(__dirname,'public','board.html')));
app.get('/admin', (req,res)=>res.sendFile(path.join(__dirname,'public','admin.html')));

app.get('/template.xlsx', (req,res)=>{
  const aoa=[
    ['Question','Option A','Option B','Option C','Option D','Bonne réponse','Image (facultatif)'],
    ['Quel est le plus grand commandement ?','Aimer Dieu et son prochain','Jeûner','Donner la dîme','Prier le matin','A',''],
    ['Qui a été restauré après avoir renié Jésus ?','Judas','Pierre','Thomas','Jean','B','']
  ];
  const ws=XLSX.utils.aoa_to_sheet(aoa); ws['!cols']=[{wch:44},{wch:22},{wch:16},{wch:16},{wch:16},{wch:16},{wch:20}];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Questions');
  const buf=XLSX.write(wb,{type:'buffer',bookType:'xlsx'});
  res.setHeader('Content-Disposition','attachment; filename="modele_questions.xlsx"');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});
app.post('/import', express.raw({type:'*/*', limit:'15mb'}), (req,res)=>{
  try{
    const name=String(req.query.name||'').toLowerCase();
    let wb; if(name.endsWith('.csv')) wb=XLSX.read(req.body.toString('utf8'),{type:'string'}); else wb=XLSX.read(req.body,{type:'buffer'});
    const qs=workbookToQuiz(wb);
    if(!qs.length) return res.status(400).json({error:'Aucune question détectée (utilise le modèle).'});
    QUESTIONS=qs.slice(0,50); if(req.query.title) QUIZ_TITLE=String(req.query.title);
    for(const id in players) delete players[id]; broadcastBoard();
    res.json({count:QUESTIONS.length, title:QUIZ_TITLE});
  }catch(e){ res.status(400).json({error:'Fichier illisible : '+e.message}); }
});

app.get('/status', (req,res)=>res.json({ count:QUESTIONS.length, title:QUIZ_TITLE, source:SOURCE, players:Object.keys(players).length, hasUrl:!!(process.env.QUESTIONS_URL||'').trim() }));
app.post('/reload', async (req,res)=>{ const r=await loadQuestions(); for(const id in players) delete players[id]; broadcastBoard(); res.json({ count:r.count, source:r.source, title:QUIZ_TITLE }); });
app.post('/reset', (req,res)=>{ for(const id in players) delete players[id]; broadcastBoard(); res.json({ ok:true }); });

/* ---------- Jeu (chacun a son rythme) ---------- */
const players={}; // socketId -> {name, score, finished}
function boardArr(){ return Object.entries(players).map(([id,p])=>({id,name:p.name,score:p.score,finished:p.finished})).sort((a,b)=>b.score-a.score); }
function broadcastBoard(){ io.to('board').emit('leaderboard',{ title:QUIZ_TITLE, board:boardArr().slice(0,20), count:Object.keys(players).length }); }

io.on('connection', (socket)=>{
  socket.on('board:join', ()=>{ socket.join('board'); socket.emit('leaderboard',{ title:QUIZ_TITLE, board:boardArr().slice(0,20), count:Object.keys(players).length }); });

  socket.on('join', ({name})=>{
    name=(name||'').trim().slice(0,24)||'Joueur';
    players[socket.id]={ name, score:0, finished:false };
    // Questions SANS la bonne reponse (anti-triche)
    const qs=QUESTIONS.map(q=>({ q:q.q, o:q.o, img:q.img||'' }));
    socket.emit('start', { title:QUIZ_TITLE, total:qs.length, questions:qs });
    broadcastBoard();
  });

  socket.on('answer', ({index, choice, timeMs})=>{
    const p=players[socket.id]; if(!p) return;
    const q=QUESTIONS[index]; if(!q) return;
    const correct=(choice===q.a);
    let gain=0;
    if(correct){ const t=Math.max(0,Math.min(20,(timeMs||0)/1000)); gain=500+Math.round(500*(20-t)/20); p.score+=gain; }
    socket.emit('answerResult', { index, correct, correctIdx:q.a, gain, score:p.score });
    broadcastBoard();
  });

  socket.on('finish', ()=>{
    const p=players[socket.id]; if(!p) return; p.finished=true;
    const arr=boardArr(); const rank=arr.findIndex(x=>x.id===socket.id)+1;
    socket.emit('final', { score:p.score, rank:rank||arr.length, total:arr.length, podium:arr.slice(0,3) });
    broadcastBoard();
  });

  socket.on('disconnect', ()=>{
    const p=players[socket.id];
    // On garde au classement ceux qui ont fini ou marque des points
    // (pour que le podium reste affiche meme s'ils ferment leur telephone).
    if(p && !p.finished && p.score===0){ delete players[socket.id]; broadcastBoard(); }
  });
});

loadQuestions().then((r)=>{
  server.listen(PORT, ()=>{
    console.log('\n==================================================');
    console.log('  QUIZ (tout sur le telephone) DEMARRE');
    console.log('  Jouer            : http://localhost:'+PORT+'/');
    console.log('  Classement ecran : http://localhost:'+PORT+'/board');
    console.log('  Admin (import)   : http://localhost:'+PORT+'/admin');
    console.log('  '+r.count+' questions ('+r.source+') : '+QUIZ_TITLE);
    console.log('==================================================\n');
  });
});
